using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Windows.Forms.VisualStyles;
using WindowInspector.Models;
using WindowInspector.Services;
using WindowInspector.Utils;
using Microsoft.VisualBasic;

namespace WindowInspector
{
    public partial class MainForm : Form
    {
        private readonly ConfigManager _configManager;
        private readonly WindowSelector _windowSelector;
        private readonly InputRecorder _inputRecorder;
        private readonly TextFiller _textFiller;
        private readonly ExcelService _excelService;
        private readonly MoleHunter _moleHunter;
        private readonly ThemeManager _themeManager;
        
        private WindowConfig _config;
        private IntPtr _targetWindow;
        private WindowHelper.RECT _windowRect;
        private CancellationTokenSource? _recordingCts;
        private List<InputPosition> _backupPositions = new();
        private System.Windows.Forms.Timer? _capsLockTimer;
        private string? _currentConfigName;
        private string? _lastExcelPath;
        
        private List<MoleGroup> _moleGroups = new();
        private string _molesDirectory;
        private int _currentMoleGroupIndex = 0;
        private int _batchSelectSliderA = 1; // 保存滑块 A 的位置
        private int _batchSelectSliderB = 1; // 保存滑块 B 的位置
        private Form? _currentEditDialog = null; // 当前打开的编辑窗口
        
        private const int HOTKEY_ID_F2 = 1;
        private const int HOTKEY_ID_F3 = 2;
        private const int HOTKEY_ID_F4 = 3;
        private const int HOTKEY_ID_F6 = 4;

        public MainForm()
        {
            InitializeComponent();
            _configManager = new ConfigManager();
            _windowSelector = new WindowSelector();
            _inputRecorder = new InputRecorder();
            _textFiller = new TextFiller();
            _excelService = new ExcelService();
            _moleHunter = new MoleHunter();
            _themeManager = new ThemeManager(_configManager);
            _config = new WindowConfig();
            
            // 初始化地鼠目录（保存到AppData）
            _molesDirectory = Path.Combine(_configManager.ProgramDirectory, "moles");
            if (!Directory.Exists(_molesDirectory))
                Directory.CreateDirectory(_molesDirectory);
            
            SetupEventHandlers();
            LoadConfiguration();
            LoadLastExcelPath();
            LoadMoles();
            ProcessPendingDeletions(); // 处理上次未能删除的文件
            RegisterGlobalHotKeys();
            
            // 应用主题
            _themeManager.ApplyTheme(this);
            ApplyTitleBarTheme();
        }

        private void ApplyTitleBarTheme()
        {
            var effectiveTheme = _themeManager.GetEffectiveTheme();
            if (effectiveTheme == ThemeMode.Dark)
            {
                WindowHelper.UseImmersiveDarkMode(this.Handle, true);
            }
            else
            {
                WindowHelper.UseImmersiveDarkMode(this.Handle, false);
            }
        }

        private void LoadConfiguration()
        {
            // 尝试加载上次使用的配置
            var lastConfigName = _configManager.LoadLastConfig();
            if (!string.IsNullOrEmpty(lastConfigName))
            {
                var configPath = Path.Combine(_configManager.ConfigsDirectory, lastConfigName + ".json");
                if (File.Exists(configPath))
                {
                    try
                    {
                        var json = File.ReadAllText(configPath);
                        var config = Newtonsoft.Json.JsonConvert.DeserializeObject<WindowConfig>(json);
                        if (config != null)
                        {
                            _config = config;
                            _currentConfigName = lastConfigName;
                            UpdateTextCombo();
                            UpdateCellGroupCombo();
                            TryAutoFindWindow();
                            UpdateWindowTitle();
                        }
                    }
                    catch { }
                }
            }
            else
            {
                // 加载默认配置
                var config = _configManager.LoadConfig();
                if (config != null)
                {
                    _config = config;
                    UpdateTextCombo();
                    UpdateCellGroupCombo();
                    TryAutoFindWindow();
                }
            }

            var windowPos = _configManager.LoadWindowPosition();
            if (windowPos != null)
            {
                // 优先使用保存的尺寸
                if (windowPos.Width > 0 && windowPos.Height > 0)
                {
                    Size = new System.Drawing.Size(windowPos.Width, windowPos.Height);
                }
                
                // 如果位置有效，使用保存的位置
                if (windowPos.X > 0 && windowPos.Y > 0)
                {
                    StartPosition = FormStartPosition.Manual;
                    Location = new System.Drawing.Point(windowPos.X, windowPos.Y);
                }
            }
        }

        private void SetupEventHandlers()
        {
            FormClosing += MainForm_FormClosing;
            
            _windowSelector.WindowSelected += WindowSelector_WindowSelected;
            _windowSelector.SelectionTimeout += (s, msg) => AppendLog(msg, LogType.Warning);
            
            _inputRecorder.InputRecorded += InputRecorder_InputRecorded;
            _inputRecorder.RecordingMessage += (s, msg) => AppendLog(msg);
            _inputRecorder.RecordingCancelled += InputRecorder_RecordingCancelled;
            _inputRecorder.RecordingCompleted += InputRecorder_RecordingCompleted;
            
            cmbSavedTexts.SelectedIndexChanged += CmbSavedTexts_SelectedIndexChanged;
            cmbCellGroups.SelectedIndexChanged += CmbCellGroups_SelectedIndexChanged;
            
            // 设置下拉框自定义绘制
            SetupComboBoxDrawing();
            
            // 启动Caps Lock监控
            StartCapsLockMonitor();
            
            // 设置文本下拉框右键菜单
            SetupTextComboContextMenu();
            
            // 设置打地鼠事件
            _moleHunter.LogMessage += (s, msg) => AppendLog(msg);
            _moleHunter.MoleFound += (s, e) => AppendLog($"🎯 击中地鼠: {e.MoleName} at ({e.Location.X}, {e.Location.Y})", LogType.Success);
            _moleHunter.HuntingStopped += MoleHunter_HuntingStopped;
        }

        private void SetupComboBoxDrawing()
        {
            cmbSavedTexts.DrawMode = DrawMode.OwnerDrawFixed;
            cmbSavedTexts.DrawItem += CmbSavedTexts_DrawItem;
        }

        private void CmbSavedTexts_DrawItem(object? sender, DrawItemEventArgs e)
        {
            if (e.Index < 0) return;

            e.DrawBackground();

            bool capsLockOn = Control.IsKeyLocked(Keys.CapsLock);
            
            // 根据主题获取正确的文字颜色
            var effectiveTheme = _themeManager.GetEffectiveTheme();
            var defaultTextColor = effectiveTheme == ThemeMode.Dark 
                ? Color.FromArgb(240, 240, 240) 
                : SystemColors.WindowText;
            
            var textColor = capsLockOn ? Color.Red : defaultTextColor;

            using (var brush = new SolidBrush(textColor))
            {
                e.Graphics.DrawString(
                    cmbSavedTexts.Items[e.Index].ToString(),
                    e.Font ?? cmbSavedTexts.Font,
                    brush,
                    e.Bounds);
            }

            e.DrawFocusRectangle();
        }

        private void RegisterGlobalHotKeys()
        {
            // 注册F2为全局热键（无修饰符）
            bool success = WindowHelper.RegisterHotKey(this.Handle, HOTKEY_ID_F2, WindowHelper.MOD_NONE, WindowHelper.VK_F2);
            if (!success)
            {
                AppendLog("⚠️ 注册F2全局热键失败，可能已被其他程序占用", LogType.Warning);
            }
            
            // 注册F3为全局热键（无修饰符）
            success = WindowHelper.RegisterHotKey(this.Handle, HOTKEY_ID_F3, WindowHelper.MOD_NONE, WindowHelper.VK_F3);
            if (!success)
            {
                AppendLog("⚠️ 注册F3全局热键失败，可能已被其他程序占用", LogType.Warning);
            }
            
            // 注册F4为全局热键（无修饰符）
            success = WindowHelper.RegisterHotKey(this.Handle, HOTKEY_ID_F4, WindowHelper.MOD_NONE, WindowHelper.VK_F4);
            if (!success)
            {
                AppendLog("⚠️ 注册F4全局热键失败，可能已被其他程序占用", LogType.Warning);
            }
            
            // 注册F6为全局热键（无修饰符）
            success = WindowHelper.RegisterHotKey(this.Handle, HOTKEY_ID_F6, WindowHelper.MOD_NONE, WindowHelper.VK_F6);
            if (!success)
            {
                AppendLog("⚠️ 注册F6全局热键失败，可能已被其他程序占用", LogType.Warning);
            }
        }

        private void UnregisterGlobalHotKeys()
        {
            WindowHelper.UnregisterHotKey(this.Handle, HOTKEY_ID_F2);
            WindowHelper.UnregisterHotKey(this.Handle, HOTKEY_ID_F3);
            WindowHelper.UnregisterHotKey(this.Handle, HOTKEY_ID_F4);
            WindowHelper.UnregisterHotKey(this.Handle, HOTKEY_ID_F6);
        }

        protected override void WndProc(ref Message m)
        {
            // 处理全局热键消息
            if (m.Msg == WindowHelper.WM_HOTKEY)
            {
                int hotkeyId = m.WParam.ToInt32();
                if (hotkeyId == HOTKEY_ID_F2)
                {
                    // F2热键被触发，执行填充操作
                    BtnFillText_Click(null, EventArgs.Empty);
                }
                else if (hotkeyId == HOTKEY_ID_F3)
                {
                    // F3热键被触发，切换打地鼠状态
                    bool isCurrentlyRunning = chkMoleEnabled.Checked;
                    
                    if (!isCurrentlyRunning)
                    {
                        // 当前未运行，即将启动 - 切换到文本填充界面
                        tabMain.SelectedIndex = 0;
                    }
                    else
                    {
                        // 当前正在运行，即将停止 - 切换到打地鼠界面
                        tabMain.SelectedIndex = 1;
                    }
                    
                    chkMoleEnabled.Checked = !chkMoleEnabled.Checked;
                }
                else if (hotkeyId == HOTKEY_ID_F4)
                {
                    // F4热键被触发，截图创建地鼠
                    BtnCaptureMole_Click(null, EventArgs.Empty);
                }
                else if (hotkeyId == HOTKEY_ID_F6)
                {
                    // F6热键被触发，添加空击位置
                    BtnSetIdleClick_Click(null, EventArgs.Empty);
                }
            }
            
            base.WndProc(ref m);
        }

        private void SetupTextComboContextMenu()
        {
            var contextMenu = new ContextMenuStrip();
            
            // 动态菜单，根据选中项的类型显示不同选项
            contextMenu.Opening += (s, e) =>
            {
                contextMenu.Items.Clear();
                
                if (cmbSavedTexts.SelectedIndex < 0 || cmbSavedTexts.SelectedIndex >= _config.SavedTexts.Count)
                {
                    e.Cancel = true;
                    return;
                }
                
                var selectedItem = _config.SavedTexts[cmbSavedTexts.SelectedIndex];
                
                // 删除选项
                var deleteItem = new ToolStripMenuItem("删除此条数据");
                deleteItem.Click += (sender, args) => DeleteSelectedText();
                contextMenu.Items.Add(deleteItem);
                
                // 重命名选项
                var renameItem = new ToolStripMenuItem("重命名");
                renameItem.Click += (sender, args) => RenameSelectedText();
                contextMenu.Items.Add(renameItem);
                
                contextMenu.Items.Add(new ToolStripSeparator());
                
                // 如果是Excel数据，显示固化选项
                if (selectedItem.FromExcel)
                {
                    var solidifyItem = new ToolStripMenuItem("固化此条数据");
                    solidifyItem.Click += (sender, args) => SolidifySingleItem();
                    contextMenu.Items.Add(solidifyItem);
                    
                    contextMenu.Items.Add(new ToolStripSeparator());
                }
                
                // 批量操作
                var deleteAllExcelItem = new ToolStripMenuItem("删除所有Excel数据");
                deleteAllExcelItem.Click += (sender, args) => DeleteAllExcelData();
                contextMenu.Items.Add(deleteAllExcelItem);
                
                var solidifyAllItem = new ToolStripMenuItem("固化所有Excel数据");
                solidifyAllItem.Click += (sender, args) => SaveExcelDataToConfig();
                contextMenu.Items.Add(solidifyAllItem);
            };
            
            cmbSavedTexts.ContextMenuStrip = contextMenu;
        }

        private void DeleteSelectedText()
        {
            if (cmbSavedTexts.SelectedIndex < 0)
                return;

            var result = MessageBox.Show(
                "确定要删除这条记录吗?",
                "确认删除",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question);

            if (result == DialogResult.Yes)
            {
                _config.SavedTexts.RemoveAt(cmbSavedTexts.SelectedIndex);
                UpdateTextCombo();
                SaveCurrentConfig();
                AppendLog("✅ 已删除记录", LogType.Success);
            }
        }

        private void RenameSelectedText()
        {
            if (cmbSavedTexts.SelectedIndex < 0)
                return;

            var item = _config.SavedTexts[cmbSavedTexts.SelectedIndex];
            var dialog = new Form
            {
                Text = "重命名",
                Size = new Size(350, 150),
                StartPosition = FormStartPosition.CenterParent,
                FormBorderStyle = FormBorderStyle.FixedDialog,
                MaximizeBox = false,
                MinimizeBox = false
            };

            var label = new Label
            {
                Text = "请输入新名称:",
                Location = new Point(20, 20),
                Size = new Size(300, 20),
                Parent = dialog
            };

            var textBox = new TextBox
            {
                Text = item.Name,
                Location = new Point(20, 45),
                Size = new Size(300, 25),
                Parent = dialog
            };

            var btnOk = new Button
            {
                Text = "确定",
                Location = new Point(150, 80),
                Size = new Size(80, 30),
                DialogResult = DialogResult.OK,
                Parent = dialog
            };

            var btnCancel = new Button
            {
                Text = "取消",
                Location = new Point(240, 80),
                Size = new Size(80, 30),
                DialogResult = DialogResult.Cancel,
                Parent = dialog
            };

            dialog.AcceptButton = btnOk;
            dialog.CancelButton = btnCancel;

            if (dialog.ShowDialog() == DialogResult.OK && !string.IsNullOrWhiteSpace(textBox.Text))
            {
                item.Name = textBox.Text.Trim();
                UpdateTextCombo();
                SaveCurrentConfig();
                AppendLog("✅ 已重命名", LogType.Success);
            }
        }

        private void BtnConfigOps_MouseDown(object? sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                ShowConfigDeleteMenu();
            }
            else if (e.Button == MouseButtons.Right)
            {
                ShowConfigLoadMenu();
            }
        }

        private void ShowConfigDeleteMenu()
        {
            var menu = new ContextMenuStrip();
            
            var themeItem = new ToolStripMenuItem("主题设置...");
            themeItem.Click += (s, e) => ShowThemeSettings();
            menu.Items.Add(themeItem);
            
            var openConfigFolderItem = new ToolStripMenuItem("打开配置文件夹");
            openConfigFolderItem.Click += (s, e) => OpenConfigFolder();
            menu.Items.Add(openConfigFolderItem);
            
            menu.Items.Add(new ToolStripSeparator());
            
            var saveAsItem = new ToolStripMenuItem("另存为配置...");
            saveAsItem.Click += (s, e) => SaveConfigAs();
            menu.Items.Add(saveAsItem);
            
            menu.Items.Add(new ToolStripSeparator());
            
            var resetFillStatusItem = new ToolStripMenuItem("重置所有填充状态");
            resetFillStatusItem.Click += (s, e) =>
            {
                var result = MessageBox.Show(
                    "确定要重置所有文本项的填充状态吗？",
                    "确认重置",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Question);

                if (result == DialogResult.Yes)
                {
                    foreach (var item in _config.SavedTexts)
                    {
                        item.LastFilledTime = null;
                    }
                    SaveCurrentConfig();
                    AppendLog("✅ 已重置所有填充状态", LogType.Success);
                    ShowAllTextItemsStatus();
                }
            };
            menu.Items.Add(resetFillStatusItem);
            
            menu.Items.Add(new ToolStripSeparator());
            
            var clearItem = new ToolStripMenuItem("清除当前配置");
            clearItem.Click += (s, e) =>
            {
                var result = MessageBox.Show(
                    "确定要清除当前配置吗？这将删除所有保存的文本和位置信息，并删除配置文件。",
                    "确认清除",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Warning);

                if (result == DialogResult.Yes)
                {
                    // 如果有命名配置，删除对应的配置文件
                    if (!string.IsNullOrEmpty(_currentConfigName))
                    {
                        var configPath = Path.Combine(_configManager.ConfigsDirectory, _currentConfigName + ".json");
                        try
                        {
                            if (File.Exists(configPath))
                            {
                                File.Delete(configPath);
                                AppendLog($"✅ 已删除配置文件: {_currentConfigName}", LogType.Success);
                            }
                        }
                        catch (Exception ex)
                        {
                            AppendLog($"⚠️ 删除配置文件失败: {ex.Message}", LogType.Warning);
                        }
                    }
                    
                    _config = new WindowConfig();
                    _targetWindow = IntPtr.Zero;
                    _currentConfigName = null;
                    UpdateTextCombo();
                    UpdateCellGroupCombo();
                    SaveCurrentConfig();
                    AppendLog("✅ 配置已清除", LogType.Success);
                    btnRecordInput.Enabled = false;
                    UpdateWindowTitle();
                }
            };
            menu.Items.Add(clearItem);
            
            menu.Show(btnConfigOps, new Point(0, btnConfigOps.Height));
        }

        private void ShowConfigLoadMenu()
        {
            var menu = new ContextMenuStrip();
            
            var loadItem = new ToolStripMenuItem("加载配置...");
            loadItem.Click += (s, e) => LoadConfigFromFile();
            menu.Items.Add(loadItem);
            
            menu.Items.Add(new ToolStripSeparator());
            
            // 列出configs目录下的所有配置文件
            var configsDir = _configManager.ConfigsDirectory;
            if (Directory.Exists(configsDir))
            {
                var configFiles = Directory.GetFiles(configsDir, "*.json");
                if (configFiles.Length > 0)
                {
                    foreach (var configFile in configFiles)
                    {
                        var fileName = Path.GetFileNameWithoutExtension(configFile);
                        var configItem = new ToolStripMenuItem(fileName);
                        configItem.Click += (s, e) => LoadNamedConfig(fileName);
                        menu.Items.Add(configItem);
                    }
                }
                else
                {
                    var noConfigItem = new ToolStripMenuItem("(无保存的配置)");
                    noConfigItem.Enabled = false;
                    menu.Items.Add(noConfigItem);
                }
            }
            
            menu.Show(btnConfigOps, new Point(0, btnConfigOps.Height));
        }

        private void SaveConfigAs()
        {
            if (_config.InputPositions.Count == 0 && _config.ExcelCells.Count == 0)
            {
                MessageBox.Show("当前没有可保存的配置", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            var dialog = new Form
            {
                Text = "另存为配置",
                Size = new Size(400, 180),
                StartPosition = FormStartPosition.CenterParent,
                FormBorderStyle = FormBorderStyle.FixedDialog,
                MaximizeBox = false,
                MinimizeBox = false
            };

            var label = new Label
            {
                Text = "请输入配置名称:",
                Location = new Point(20, 20),
                Size = new Size(350, 20),
                Parent = dialog
            };

            var textBox = new TextBox
            {
                Text = _currentConfigName ?? (_config.WindowTitle ?? "新配置"),
                Location = new Point(20, 45),
                Size = new Size(350, 25),
                Parent = dialog
            };

            var hintLabel = new Label
            {
                Text = "提示：配置将保存到 configs 目录",
                Location = new Point(20, 75),
                Size = new Size(350, 20),
                ForeColor = Color.Gray,
                Parent = dialog
            };

            var btnOk = new Button
            {
                Text = "保存",
                Location = new Point(200, 110),
                Size = new Size(80, 30),
                DialogResult = DialogResult.OK,
                Parent = dialog
            };

            var btnCancel = new Button
            {
                Text = "取消",
                Location = new Point(290, 110),
                Size = new Size(80, 30),
                DialogResult = DialogResult.Cancel,
                Parent = dialog
            };

            dialog.AcceptButton = btnOk;
            dialog.CancelButton = btnCancel;

            if (dialog.ShowDialog() == DialogResult.OK && !string.IsNullOrWhiteSpace(textBox.Text))
            {
                var configName = textBox.Text.Trim();
                var configPath = Path.Combine(_configManager.ConfigsDirectory, configName + ".json");
                
                try
                {
                    var json = Newtonsoft.Json.JsonConvert.SerializeObject(_config, Newtonsoft.Json.Formatting.Indented);
                    File.WriteAllText(configPath, json);
                    _currentConfigName = configName;
                    _configManager.SaveLastConfig(configName);
                    AppendLog($"✅ 配置已保存为: {configName}", LogType.Success);
                    UpdateWindowTitle();
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"保存配置失败: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void LoadConfigFromFile()
        {
            var ofd = new OpenFileDialog
            {
                Filter = "配置文件|*.json",
                Title = "选择配置文件",
                InitialDirectory = _configManager.ConfigsDirectory
            };

            if (ofd.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    var json = File.ReadAllText(ofd.FileName);
                    var config = Newtonsoft.Json.JsonConvert.DeserializeObject<WindowConfig>(json);
                    if (config != null)
                    {
                        _config = config;
                        _currentConfigName = Path.GetFileNameWithoutExtension(ofd.FileName);
                        UpdateTextCombo();
                        UpdateCellGroupCombo();
                        TryAutoFindWindow();
                        _configManager.SaveLastConfig(_currentConfigName);
                        AppendLog($"✅ 已加载配置: {_currentConfigName}", LogType.Success);
                        UpdateWindowTitle();
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"加载配置失败: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void LoadNamedConfig(string configName)
        {
            var configPath = Path.Combine(_configManager.ConfigsDirectory, configName + ".json");

            try
            {
                var json = File.ReadAllText(configPath);
                var config = Newtonsoft.Json.JsonConvert.DeserializeObject<WindowConfig>(json);
                if (config != null)
                {
                    _config = config;
                    _currentConfigName = configName;
                    UpdateTextCombo();
                    UpdateCellGroupCombo();
                    TryAutoFindWindow();
                    _configManager.SaveLastConfig(configName);
                    AppendLog($"✅ 已加载配置: {configName}", LogType.Success);
                    UpdateWindowTitle();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"加载配置失败: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void UpdateWindowTitle()
        {
            if (!string.IsNullOrEmpty(_currentConfigName))
            {
                Text = $"文本框位置记录工具 - [{_currentConfigName}]";
            }
            else
            {
                Text = "文本框位置记录工具";
            }
        }

        private void ShowThemeSettings()
        {
            using (var dialog = new Dialogs.ThemeSettingsDialog(_themeManager))
            {
                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    _themeManager.ChangeTheme(dialog.SelectedTheme);
                    _themeManager.ApplyTheme(this);
                    ApplyTitleBarTheme();
                    AppendLog($"✅ 主题已切换为: {GetThemeModeName(dialog.SelectedTheme)}", LogType.Success);
                }
            }
        }

        private string GetThemeModeName(ThemeMode mode)
        {
            return mode switch
            {
                ThemeMode.Light => "浅色主题",
                ThemeMode.Dark => "深色主题",
                ThemeMode.System => "随系统",
                _ => "未知"
            };
        }

        private void OpenConfigFolder()
        {
            try
            {
                var configPath = _configManager.ProgramDirectory;
                if (Directory.Exists(configPath))
                {
                    System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo()
                    {
                        FileName = configPath,
                        UseShellExecute = true,
                        Verb = "open"
                    });
                    AppendLog($"📁 已打开配置文件夹: {configPath}", LogType.Info);
                }
                else
                {
                    MessageBox.Show("配置文件夹不存在", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"打开文件夹失败: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void StartCapsLockMonitor()
        {
            _capsLockTimer = new System.Windows.Forms.Timer();
            _capsLockTimer.Interval = 100; // 每100ms检查一次
            _capsLockTimer.Tick += (s, e) =>
            {
                bool capsLockOn = Control.IsKeyLocked(Keys.CapsLock);
                pnlCapsIndicator.BackColor = capsLockOn ? Color.Red : Color.Green;
                cmbSavedTexts.ForeColor = capsLockOn ? Color.Red : SystemColors.WindowText;
            };
            _capsLockTimer.Start();
        }

        private void MainForm_FormClosing(object? sender, FormClosingEventArgs e)
        {
            _capsLockTimer?.Stop();
            _capsLockTimer?.Dispose();
            
            // 清理预览窗口
            HidePreview();
            if (_previewForm != null)
            {
                _previewForm.Dispose();
                _previewForm = null;
            }
            
            // 注销全局热键
            UnregisterGlobalHotKeys();
            
            var windowPos = new WindowPosition
            {
                X = Location.X,
                Y = Location.Y,
                Width = Width,
                Height = Height
            };
            _configManager.SaveWindowPosition(windowPos);
        }

        private void TryAutoFindWindow()
        {
            if (string.IsNullOrEmpty(_config.WindowClass))
                return;

            IntPtr foundWindow = IntPtr.Zero;
            WindowHelper.EnumWindows((hwnd, lParam) =>
            {
                var className = WindowHelper.GetWindowClassName(hwnd);
                var title = WindowHelper.GetWindowTitle(hwnd);

                if (_config.IsExcelMode)
                {
                    if (className == _config.WindowClass)
                    {
                        foundWindow = hwnd;
                        return false;
                    }
                }
                else
                {
                    if (className == _config.WindowClass && title == _config.WindowTitle)
                    {
                        foundWindow = hwnd;
                        return false;
                    }
                }
                return true;
            }, IntPtr.Zero);

            if (foundWindow != IntPtr.Zero)
            {
                _targetWindow = foundWindow;
                WindowHelper.GetWindowRect(_targetWindow, out _windowRect);
                OnWindowSelected(_config.WindowTitle, true);
            }
        }

        private void OnWindowSelected(string windowTitle, bool auto)
        {
            var source = auto ? "自动加载" : "已选择";
            AppendLog($"\n{source}窗口: {windowTitle}");

            if (_config.IsExcelMode)
            {
                AppendLog("📊 检测到Excel窗口，已切换到Excel专用模式", LogType.Success);
            }
            else
            {
                AppendLog("📝 普通窗口模式", LogType.Info);
            }

            if (_config.InputPositions.Count > 0)
            {
                AppendLog("\n已加载输入框位置:");
                for (int i = 0; i < _config.InputPositions.Count; i++)
                {
                    var pos = _config.InputPositions[i];
                    AppendLog($"输入框 {i + 1}: 相对位置 ({pos.X}, {pos.Y})");
                }
            }
        }

        private void UpdateTextCombo()
        {
            cmbSavedTexts.Items.Clear();
            foreach (var item in _config.SavedTexts)
            {
                var displayName = item.FromExcel ? $"📊 {item.Name}" : item.Name;
                cmbSavedTexts.Items.Add(displayName);
            }
            if (cmbSavedTexts.Items.Count > 0)
                cmbSavedTexts.SelectedIndex = 0;
        }

        private void UpdateCellGroupCombo()
        {
            cmbCellGroups.Items.Clear();
            foreach (var group in _config.ExcelCellGroups)
            {
                cmbCellGroups.Items.Add(group.Name);
            }
            if (_config.ActiveCellGroupIndex < cmbCellGroups.Items.Count)
                cmbCellGroups.SelectedIndex = _config.ActiveCellGroupIndex;
        }

        private void AppendLog(string message, LogType type = LogType.Normal)
        {
            if (InvokeRequired)
            {
                Invoke(new Action(() => AppendLog(message, type)));
                return;
            }

            var effectiveTheme = _themeManager.GetEffectiveTheme();
            
            // 根据主题选择颜色
            var color = effectiveTheme == ThemeMode.Dark ? 
                type switch
                {
                    LogType.Success => Color.FromArgb(76, 175, 80),      // 绿色
                    LogType.Warning => Color.FromArgb(255, 152, 0),      // 橙色
                    LogType.Error => Color.FromArgb(244, 67, 54),        // 红色
                    LogType.Info => Color.FromArgb(33, 150, 243),        // 蓝色
                    _ => Color.White
                }
                :
                type switch
                {
                    LogType.Success => Color.FromArgb(56, 142, 60),      // 深绿色
                    LogType.Warning => Color.FromArgb(230, 124, 0),      // 深橙色
                    LogType.Error => Color.FromArgb(211, 47, 47),        // 深红色
                    LogType.Info => Color.FromArgb(13, 71, 161),         // 深蓝色
                    _ => Color.FromArgb(30, 30, 30)                      // 深灰色
                };

            rtbLog.SelectionStart = rtbLog.TextLength;
            rtbLog.SelectionLength = 0;
            rtbLog.SelectionColor = color;
            rtbLog.AppendText(message + "\n");
            rtbLog.SelectionColor = rtbLog.ForeColor;
            rtbLog.ScrollToCaret();
        }

        private void WindowSelector_WindowSelected(object? sender, WindowSelectedEventArgs e)
        {
            _targetWindow = e.WindowHandle;
            _windowRect = e.WindowRect;
            
            _config.WindowClass = e.WindowClass;
            _config.WindowTitle = e.WindowTitle;
            _config.IsExcelMode = WindowHelper.IsExcelWindow(_targetWindow);
            
            var programPath = WindowHelper.GetProcessPath(_targetWindow);
            if (!string.IsNullOrEmpty(programPath))
            {
                _config.TargetProgramPath = programPath;
                var result = MessageBox.Show(
                    $"是否在找不到目标窗口时自动启动程序?\n路径: {programPath}",
                    "自动启动",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Question);
                _config.AutoLaunch = result == DialogResult.Yes;
            }
            
            OnWindowSelected(e.WindowTitle, false);
            btnRecordInput.Enabled = true;
        }

        private void InputRecorder_InputRecorded(object? sender, InputRecordedEventArgs e)
        {
            AppendLog($"✅ 已记录第 {e.Index + 1} 个位置: ({e.Position.X}, {e.Position.Y})", LogType.Success);
        }

        private void InputRecorder_RecordingCancelled(object? sender, EventArgs e)
        {
            _config.InputPositions = _backupPositions;
            AppendLog("\n❌ 已取消记录操作", LogType.Warning);
            btnRecordInput.Enabled = true;
            btnRecordInput.Text = "2. 记录输入框位置";
        }

        private void InputRecorder_RecordingCompleted(object? sender, List<InputPosition> positions)
        {
            _config.InputPositions = positions;
            AppendLog($"\n🎉 已完成 {positions.Count} 个输入框位置的记录", LogType.Success);
            btnRecordInput.Enabled = true;
            btnRecordInput.Text = "重新记录输入框位置";
            
            // 提示用户保存配置
            PromptSaveConfig();
        }

        private void PromptSaveConfig()
        {
            var result = MessageBox.Show(
                "是否为此配置命名并保存？\n\n点击\"是\"保存配置\n点击\"否\"仅临时使用",
                "保存配置",
                MessageBoxButtons.YesNoCancel,
                MessageBoxIcon.Question);

            if (result == DialogResult.Yes)
            {
                SaveConfigAs();
            }
            else if (result == DialogResult.No)
            {
                // 仅保存到默认配置
                SaveCurrentConfig();
                AppendLog("✅ 配置已临时保存", LogType.Success);
            }
            // Cancel 则不保存
        }

        private void CmbSavedTexts_SelectedIndexChanged(object? sender, EventArgs e)
        {
            if (cmbSavedTexts.SelectedIndex >= 0 && cmbSavedTexts.SelectedIndex < _config.SavedTexts.Count)
            {
                var item = _config.SavedTexts[cmbSavedTexts.SelectedIndex];
                
                // 显示当前选中的文本内容
                AppendLog($"\n▶️ 当前选中: {item.Name}", LogType.Info);
                for (int i = 0; i < item.Texts.Count; i++)
                {
                    // 如果只有2个文本且是第2个，用*号显示
                    if (item.Texts.Count == 2 && i == 1)
                    {
                        AppendLog($"文本{i + 1}: {new string('*', item.Texts[i].Length)}");
                    }
                    else
                    {
                        AppendLog($"文本{i + 1}: {item.Texts[i]}");
                    }
                }
                
                // 显示填充状态
                if (item.LastFilledTime.HasValue)
                {
                    var timeDiff = CalculateTimeDifference(item.LastFilledTime.Value);
                    AppendLog($"状态: ✅ 已填充 ({timeDiff})", LogType.Success);
                }
                else
                {
                    AppendLog("状态: ⏸️ 待填充", LogType.Warning);
                }
                
                AppendLog("");
            }
        }

        private void CmbCellGroups_SelectedIndexChanged(object? sender, EventArgs e)
        {
            if (cmbCellGroups.SelectedIndex >= 0 && cmbCellGroups.SelectedIndex < _config.ExcelCellGroups.Count)
            {
                _config.ActiveCellGroupIndex = cmbCellGroups.SelectedIndex;
                var group = _config.ExcelCellGroups[cmbCellGroups.SelectedIndex];
                _config.ExcelCells = group.Cells;
                txtInputCount.Text = group.Cells.Count.ToString();
                AppendLog($"\n📍 已切换到地址组: {group.Name}", LogType.Info);
                SaveCurrentConfig();
            }
        }

        // 按钮事件处理器
        internal async void BtnSelectWindow_Click(object? sender, EventArgs e)
        {
            AppendLog("\n请点击要操作的窗口...");
            btnSelectWindow.Enabled = false;
            var cts = new CancellationTokenSource();
            await _windowSelector.StartSelectionAsync(cts.Token);
            btnSelectWindow.Enabled = true;
        }

        internal async void BtnRecordInput_Click(object? sender, EventArgs e)
        {
            if (!int.TryParse(txtInputCount.Text, out int count) || count < 1)
            {
                MessageBox.Show("请输入有效的数字", "错误", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (_config.IsExcelMode)
            {
                RecordExcelCells(count);
                return;
            }

            _backupPositions = new List<InputPosition>(_config.InputPositions);
            _config.InputPositions.Clear();
            
            AppendLog($"\n📍 开始记录 {count} 个输入框位置", LogType.Info);
            AppendLog("💡 按 ESC 键可取消操作", LogType.Info);
            
            btnRecordInput.Enabled = false;
            btnRecordInput.Text = "正在记录...";
            
            _recordingCts = new CancellationTokenSource();
            await _inputRecorder.StartRecordingAsync(_targetWindow, _windowRect, count, _recordingCts.Token);
        }

        private void RecordExcelCells(int count)
        {
            var dialog = new ExcelCellInputDialog(count);
            if (dialog.ShowDialog() == DialogResult.OK)
            {
                _config.ExcelCells = dialog.Cells;
                
                if (_config.ExcelCellGroups.Count == 0)
                {
                    _config.ExcelCellGroups.Add(new CellGroup
                    {
                        Name = "地址组1",
                        Cells = new List<string>(_config.ExcelCells)
                    });
                }
                else
                {
                    _config.ExcelCellGroups[_config.ActiveCellGroupIndex].Cells = new List<string>(_config.ExcelCells);
                }
                
                UpdateCellGroupCombo();
                AppendLog("✅ Excel单元格地址已配置", LogType.Success);
                
                // 提示用户保存配置
                PromptSaveConfig();
            }
        }

        internal void BtnSaveText_Click(object? sender, EventArgs e)
        {
            if (_config.InputPositions.Count == 0 && _config.ExcelCells.Count == 0)
            {
                MessageBox.Show("请先完成窗口和输入框位置的选择", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            var inputCount = _config.IsExcelMode ? _config.ExcelCells.Count : _config.InputPositions.Count;
            var dialog = new TextInputDialog(inputCount);
            
            if (dialog.ShowDialog() == DialogResult.OK)
            {
                var item = new SavedTextItem
                {
                    Name = dialog.ItemName,
                    Texts = dialog.Texts,
                    FromExcel = false,
                    LastFilledTime = null
                };
                
                _config.SavedTexts.Add(item);
                UpdateTextCombo();
                SaveCurrentConfig();
                AppendLog("✅ 文本已保存", LogType.Success);
            }
        }

        internal void BtnSaveText_MouseDown(object? sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right)
            {
                SaveExcelDataToConfig();
            }
        }

        private void SaveExcelDataToConfig()
        {
            // 检查是否有加载的Excel数据
            var excelItems = _config.SavedTexts.Where(item => item.FromExcel).ToList();
            
            if (excelItems.Count == 0)
            {
                MessageBox.Show("当前没有加载的Excel数据", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            // 确认操作
            var result = MessageBox.Show(
                $"确定要将当前 {excelItems.Count} 条Excel数据永久保存到配置中吗？\n\n" +
                "保存后这些数据将标记为本地数据，不再显示Excel标记。",
                "保存Excel数据",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question);

            if (result != DialogResult.Yes)
                return;

            try
            {
                // 将Excel数据标记为本地数据
                foreach (var item in excelItems)
                {
                    item.FromExcel = false;
                }

                // 保存配置到正确的位置
                SaveCurrentConfig();
                
                // 更新显示
                UpdateTextCombo();
                
                AppendLog($"✅ 已将 {excelItems.Count} 条Excel数据永久保存到配置", LogType.Success);
                AppendLog("这些数据现在已成为本地配置的一部分", LogType.Info);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"保存失败: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// 固化单条Excel数据
        /// </summary>
        private void SolidifySingleItem()
        {
            if (cmbSavedTexts.SelectedIndex < 0 || cmbSavedTexts.SelectedIndex >= _config.SavedTexts.Count)
                return;

            var item = _config.SavedTexts[cmbSavedTexts.SelectedIndex];
            
            if (!item.FromExcel)
            {
                MessageBox.Show("此数据已经是固化数据", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            var result = MessageBox.Show(
                $"确定要固化数据 \"{item.Name}\" 吗？\n\n" +
                "固化后此数据将成为本地配置的一部分，不再显示Excel标记。",
                "固化数据",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question);

            if (result != DialogResult.Yes)
                return;

            try
            {
                item.FromExcel = false;
                SaveCurrentConfig();
                UpdateTextCombo();
                
                // 保持选中当前项
                if (cmbSavedTexts.SelectedIndex >= 0)
                    cmbSavedTexts.SelectedIndex = cmbSavedTexts.SelectedIndex;
                
                AppendLog($"✅ 已固化数据: {item.Name}", LogType.Success);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"固化失败: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// 删除所有Excel数据
        /// </summary>
        private void DeleteAllExcelData()
        {
            var excelItems = _config.SavedTexts.Where(item => item.FromExcel).ToList();
            
            if (excelItems.Count == 0)
            {
                MessageBox.Show("当前没有Excel数据", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            var result = MessageBox.Show(
                $"确定要删除所有 {excelItems.Count} 条Excel数据吗？\n\n" +
                "此操作不可恢复！",
                "删除Excel数据",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Warning);

            if (result != DialogResult.Yes)
                return;

            try
            {
                foreach (var item in excelItems)
                {
                    _config.SavedTexts.Remove(item);
                }
                
                SaveCurrentConfig();
                UpdateTextCombo();
                
                AppendLog($"✅ 已删除 {excelItems.Count} 条Excel数据", LogType.Success);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"删除失败: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        internal void BtnLoadExcel_Click(object? sender, EventArgs e)
        {
            LoadExcelFile();
        }

        private void LoadExcelFile(string? filePath = null)
        {
            if (filePath == null)
            {
                // 选择Excel文件
                var ofd = new OpenFileDialog
                {
                    Filter = "Excel文件|*.xlsx;*.xls",
                    Title = "选择Excel文件导入数据"
                };

                // 如果有上次的路径，设置初始目录
                if (!string.IsNullOrEmpty(_lastExcelPath) && File.Exists(_lastExcelPath))
                {
                    ofd.InitialDirectory = Path.GetDirectoryName(_lastExcelPath);
                    ofd.FileName = Path.GetFileName(_lastExcelPath);
                }

                if (ofd.ShowDialog() != DialogResult.OK)
                    return;

                filePath = ofd.FileName;
            }

            try
            {
                // 自动从Excel读取数据，A列作为名称，B列开始作为文本
                var items = _excelService.LoadFromExcelAuto(filePath);
                
                if (items.Count == 0)
                {
                    MessageBox.Show("Excel文件中没有找到有效数据\n\n提示：A列应为名称，B列开始为文本内容", 
                        "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }
                
                _config.SavedTexts.AddRange(items);
                UpdateTextCombo();
                SaveCurrentConfig();
                
                // 保存上次加载的Excel路径
                _lastExcelPath = filePath;
                SaveLastExcelPath();
                
                AppendLog($"✅ 已从Excel导入 {items.Count} 条数据", LogType.Success);
                AppendLog($"文件: {Path.GetFileName(filePath)}", LogType.Info);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void SaveLastExcelPath()
        {
            try
            {
                var lastExcelFile = Path.Combine(_configManager.ProgramDirectory, "last_excel.txt");
                File.WriteAllText(lastExcelFile, _lastExcelPath ?? "");
            }
            catch { }
        }

        private void LoadLastExcelPath()
        {
            try
            {
                var lastExcelFile = Path.Combine(_configManager.ProgramDirectory, "last_excel.txt");
                if (File.Exists(lastExcelFile))
                {
                    _lastExcelPath = File.ReadAllText(lastExcelFile);
                }
            }
            catch { }
        }

        /// <summary>
        /// 保存当前配置到正确的位置（命名配置或默认配置）
        /// </summary>
        private void SaveCurrentConfig()
        {
            try
            {
                if (!string.IsNullOrEmpty(_currentConfigName))
                {
                    // 如果有命名配置，保存到 configs 目录
                    var configPath = Path.Combine(_configManager.ConfigsDirectory, _currentConfigName + ".json");
                    var json = Newtonsoft.Json.JsonConvert.SerializeObject(_config, Newtonsoft.Json.Formatting.Indented);
                    File.WriteAllText(configPath, json);
                }
                else
                {
                    // 否则保存到默认配置
                    _configManager.SaveConfig(_config);
                }
            }
            catch (Exception ex)
            {
                throw new Exception($"保存配置失败: {ex.Message}");
            }
        }

        internal void BtnLoadExcel_MouseDown(object? sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right)
            {
                // 如果有上次加载的Excel文件，直接自动加载
                if (!string.IsNullOrEmpty(_lastExcelPath) && File.Exists(_lastExcelPath))
                {
                    AppendLog($"\n📂 自动加载上次的Excel文件...", LogType.Info);
                    LoadExcelFile(_lastExcelPath);
                }
                else
                {
                    // 没有历史记录时显示菜单
                    ShowLoadExcelMenu();
                }
            }
        }

        private void ShowLoadExcelMenu()
        {
            var menu = new ContextMenuStrip();
            
            if (!string.IsNullOrEmpty(_lastExcelPath) && File.Exists(_lastExcelPath))
            {
                var fileName = Path.GetFileName(_lastExcelPath);
                var reloadItem = new ToolStripMenuItem($"重新加载: {fileName}");
                reloadItem.Click += (s, e) =>
                {
                    AppendLog($"\n📂 重新加载上次的Excel文件...", LogType.Info);
                    LoadExcelFile(_lastExcelPath);
                };
                menu.Items.Add(reloadItem);
                
                menu.Items.Add(new ToolStripSeparator());
            }
            
            var browseItem = new ToolStripMenuItem("浏览选择Excel文件...");
            browseItem.Click += (s, e) => LoadExcelFile();
            menu.Items.Add(browseItem);
            
            if (string.IsNullOrEmpty(_lastExcelPath) || !File.Exists(_lastExcelPath))
            {
                var noHistoryItem = new ToolStripMenuItem("(无历史记录)");
                noHistoryItem.Enabled = false;
                menu.Items.Insert(0, noHistoryItem);
                menu.Items.Insert(1, new ToolStripSeparator());
            }
            
            menu.Show(btnLoadExcel, new Point(0, btnLoadExcel.Height));
        }

        internal void BtnOpenExcel_Click(object? sender, EventArgs e)
        {
            var ofd = new OpenFileDialog
            {
                Filter = "Excel文件|*.xlsx;*.xls",
                Title = "打开Excel文件"
            };

            if (ofd.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    _excelService.OpenExcel(ofd.FileName);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        internal async void BtnFillText_Click(object? sender, EventArgs e)
        {
            if (cmbSavedTexts.SelectedIndex < 0)
            {
                MessageBox.Show("请先选择要填充的文本", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            var currentIndex = cmbSavedTexts.SelectedIndex;
            var item = _config.SavedTexts[currentIndex];
            
            try
            {
                if (_config.IsExcelMode)
                {
                    await _textFiller.FillExcelCellsAsync(_config.ExcelCells, item.Texts);
                }
                else
                {
                    await _textFiller.FillTextAsync(_targetWindow, _windowRect, _config.InputPositions, item.Texts);
                }
                
                item.LastFilledTime = DateTime.Now;
                SaveCurrentConfig();
                
                // 标记当前项为已填充
                AppendLog($"✅ 已填充: {item.Name}", LogType.Success);
                
                // 自动切换到下一个未填充的项（从当前位置往下找）
                int nextIndex = FindNextUnfilledIndex(currentIndex + 1);
                
                if (nextIndex >= 0)
                {
                    cmbSavedTexts.SelectedIndex = nextIndex;
                    var nextItem = _config.SavedTexts[nextIndex];
                    AppendLog($"⏭️ 已切换到: {nextItem.Name}", LogType.Info);
                    
                    // 显示状态
                    ShowAllTextItemsStatus();
                }
                else
                {
                    AppendLog("🎉 所有文本已填充完成！", LogType.Success);
                    // 显示状态
                    ShowAllTextItemsStatus();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"填充失败: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// 从指定位置开始查找下一个未填充的项
        /// </summary>
        private int FindNextUnfilledIndex(int startIndex)
        {
            // 从startIndex开始往后找
            for (int i = startIndex; i < _config.SavedTexts.Count; i++)
            {
                if (!_config.SavedTexts[i].LastFilledTime.HasValue)
                {
                    return i;
                }
            }
            
            // 如果后面没有未填充的，返回-1
            return -1;
        }

        private void ShowAllTextItemsStatus()
        {
            // 统计已填充数量
            int filledCount = _config.SavedTexts.Count(item => item.LastFilledTime.HasValue);
            int totalCount = _config.SavedTexts.Count;
            
            AppendLog($"\n📊 进度: {filledCount}/{totalCount} 已完成", LogType.Info);
            
            // 显示最近填充的3条
            var recentFilled = _config.SavedTexts
                .Where(item => item.LastFilledTime.HasValue)
                .OrderByDescending(item => item.LastFilledTime)
                .Take(3)
                .ToList();
            
            if (recentFilled.Count > 0)
            {
                AppendLog("最近已填充:", LogType.Success);
                foreach (var item in recentFilled)
                {
                    var timeDiff = CalculateTimeDifference(item.LastFilledTime!.Value);
                    AppendLog($"  ✅ {item.Name} ({timeDiff})", LogType.Normal);
                }
            }
            
            // 显示下一个待填充的
            var currentIndex = cmbSavedTexts.SelectedIndex;
            if (currentIndex >= 0 && currentIndex < _config.SavedTexts.Count)
            {
                var currentItem = _config.SavedTexts[currentIndex];
                if (!currentItem.LastFilledTime.HasValue)
                {
                    AppendLog($"\n▶️ 下一个待填充: {currentItem.Name}", LogType.Warning);
                }
            }
            
            AppendLog("");
        }

        private string CalculateTimeDifference(DateTime lastTime)
        {
            var diff = DateTime.Now - lastTime;
            
            if (diff.TotalMinutes < 1)
                return "刚刚";
            else if (diff.TotalMinutes < 60)
                return $"{(int)diff.TotalMinutes}分钟前";
            else if (diff.TotalHours < 24)
                return $"{(int)diff.TotalHours}小时前";
            else
                return $"{(int)diff.TotalDays}天前";
        }

        internal void BtnExportExcel_Click(object? sender, EventArgs e)
        {
            if (_config.SavedTexts.Count == 0)
            {
                MessageBox.Show("没有可导出的数据", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            var sfd = new SaveFileDialog
            {
                Filter = "Excel文件|*.xlsx",
                Title = "导出到Excel"
            };

            if (sfd.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    var cells = _config.IsExcelMode ? _config.ExcelCells : 
                        Enumerable.Range(0, _config.InputPositions.Count).Select(i => $"{(char)('A' + i)}").ToList();
                    
                    _excelService.ExportToExcel(sfd.FileName, _config.SavedTexts, cells);
                    AppendLog("✅ 导出成功", LogType.Success);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private enum LogType
        {
            Normal,
            Success,
            Warning,
            Error,
            Info
        }

        // ==================== 打地鼠功能 ====================
        
        private MoleGroup GetCurrentMoleGroup()
        {
            if (_moleGroups.Count == 0)
            {
                _moleGroups.Add(new MoleGroup { Name = "默认" });
            }
            if (_currentMoleGroupIndex >= _moleGroups.Count)
            {
                _currentMoleGroupIndex = 0;
            }
            return _moleGroups[_currentMoleGroupIndex];
        }

        private void LoadMoles()
        {
            _moleGroups.Clear();
            tabMoleGroups.TabPages.Clear();
            
            if (!Directory.Exists(_molesDirectory))
            {
                // 创建默认组
                var defaultGroup = new MoleGroup { Name = "默认" };
                _moleGroups.Add(defaultGroup);
                CreateMoleGroupTab(defaultGroup, 0);
                return;
            }
            
            // 加载分组配置
            var groupsConfigPath = Path.Combine(_molesDirectory, "mole_groups.json");
            if (File.Exists(groupsConfigPath))
            {
                try
                {
                    var json = File.ReadAllText(groupsConfigPath);
                    var loadedGroups = Newtonsoft.Json.JsonConvert.DeserializeObject<List<MoleGroup>>(json);
                    if (loadedGroups != null && loadedGroups.Count > 0)
                    {
                        _moleGroups = loadedGroups;
                    }
                }
                catch (Exception ex)
                {
                    AppendLog($"⚠️ 加载分组配置失败: {ex.Message}", LogType.Warning);
                }
            }
            
            // 如果没有加载到分组，从旧格式迁移
            if (_moleGroups.Count == 0)
            {
                var defaultGroup = new MoleGroup { Name = "默认" };
                
                // 加载旧的阈值配置
                var configPath = Path.Combine(_molesDirectory, "moles_config.json");
                Dictionary<string, double> thresholdConfig = new Dictionary<string, double>();
                
                if (File.Exists(configPath))
                {
                    try
                    {
                        var json = File.ReadAllText(configPath);
                        thresholdConfig = Newtonsoft.Json.JsonConvert.DeserializeObject<Dictionary<string, double>>(json) 
                            ?? new Dictionary<string, double>();
                    }
                    catch { }
                }
                
                // 加载所有图片文件
                var imageFiles = Directory.GetFiles(_molesDirectory, "*.png")
                    .Concat(Directory.GetFiles(_molesDirectory, "*.jpg"))
                    .Concat(Directory.GetFiles(_molesDirectory, "*.bmp"));
                
                foreach (var file in imageFiles)
                {
                    var fileName = Path.GetFileName(file);
                    var mole = new MoleItem
                    {
                        Name = Path.GetFileNameWithoutExtension(file),
                        ImagePath = file,
                        IsEnabled = true,
                        SimilarityThreshold = thresholdConfig.ContainsKey(fileName) ? thresholdConfig[fileName] : 0.85
                    };
                    defaultGroup.Moles.Add(mole);
                }
                
                _moleGroups.Add(defaultGroup);
            }
            
            // 创建标签页
            for (int i = 0; i < _moleGroups.Count; i++)
            {
                CreateMoleGroupTab(_moleGroups[i], i);
            }
            
            // 选中第一个标签页
            if (tabMoleGroups.TabPages.Count > 0)
            {
                tabMoleGroups.SelectedIndex = 0;
                _currentMoleGroupIndex = 0;
            }
            
            UpdateIdleClickLabel();
            AppendLog($"📂 已加载 {_moleGroups.Count} 个地鼠分组", LogType.Info);
        }
        
        private void CreateMoleGroupTab(MoleGroup group, int index)
        {
            var tabPage = new TabPage(group.Name);
            tabPage.Tag = index;
            
            var lstMoles = new CheckedListBox
            {
                Location = new Point(0, 0),
                Size = new Size(tabPage.ClientSize.Width, tabPage.ClientSize.Height),
                Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right,
                CheckOnClick = true,
                DrawMode = DrawMode.OwnerDrawFixed,
                // 注意：CheckedListBox 不支持 MultiExtended 模式，只能使用 One 模式
                Parent = tabPage
            };
            
            // 加载该组的地鼠
            for (int i = 0; i < group.Moles.Count; i++)
            {
                var mole = group.Moles[i];
                string displayText;
                
                if (mole.IsIdleClick && mole.IdleClickPosition.HasValue)
                {
                    displayText = $"{i + 1}. 💤 {mole.Name}: ({mole.IdleClickPosition.Value.X}, {mole.IdleClickPosition.Value.Y})";
                }
                else if (mole.IsJump)
                {
                    displayText = $"{i + 1}. 🔗 {mole.Name}";
                }
                else
                {
                    displayText = $"{i + 1}. {mole.Name}";
                }
                
                lstMoles.Items.Add(displayText, mole.IsEnabled);
            }
            
            lstMoles.MouseDown += LstMoles_MouseDown;
            lstMoles.MouseMove += LstMoles_MouseMove;
            lstMoles.MouseLeave += LstMoles_MouseLeave;
            lstMoles.KeyDown += LstMoles_KeyDown;
            lstMoles.DrawItem += LstMoles_DrawItem;
            
            // 手动应用主题颜色
            var effectiveTheme = _themeManager.GetEffectiveTheme();
            if (effectiveTheme == ThemeMode.Dark)
            {
                lstMoles.BackColor = Color.FromArgb(45, 45, 48);
                lstMoles.ForeColor = Color.FromArgb(240, 240, 240);
            }
            else
            {
                lstMoles.BackColor = Color.White;
                lstMoles.ForeColor = Color.Black;
            }
            lstMoles.BorderStyle = BorderStyle.FixedSingle;
            
            // 标记此列表，防止主题管理器接管绘制
            lstMoles.Tag = "CustomDraw";
            
            tabMoleGroups.TabPages.Add(tabPage);
        }
        
        private CheckedListBox? GetCurrentMoleListBox()
        {
            if (tabMoleGroups.SelectedTab != null)
            {
                foreach (Control ctrl in tabMoleGroups.SelectedTab.Controls)
                {
                    if (ctrl is CheckedListBox listBox)
                    {
                        return listBox;
                    }
                }
            }
            return null;
        }
        
        private void SaveMoles()
        {
            if (!Directory.Exists(_molesDirectory))
                Directory.CreateDirectory(_molesDirectory);
            
            // 保存分组配置
            var groupsConfigPath = Path.Combine(_molesDirectory, "mole_groups.json");
            try
            {
                var json = Newtonsoft.Json.JsonConvert.SerializeObject(_moleGroups, Newtonsoft.Json.Formatting.Indented);
                File.WriteAllText(groupsConfigPath, json);
            }
            catch (Exception ex)
            {
                AppendLog($"❌ 保存分组配置失败: {ex.Message}", LogType.Error);
            }
        }
        
        private void UpdateIdleClicksInList()
        {
            var group = GetCurrentMoleGroup();
            
            // 移除旧的空击项
            for (int i = group.Moles.Count - 1; i >= 0; i--)
            {
                if (group.Moles[i].IsIdleClick)
                {
                    group.Moles.RemoveAt(i);
                }
            }
            
            // 添加新的空击项
            for (int i = 0; i < group.IdleClickPositions.Count; i++)
            {
                var pos = group.IdleClickPositions[i];
                var idleMole = new MoleItem
                {
                    Name = $"空击 {i + 1}",
                    ImagePath = "",
                    IsEnabled = true,
                    IsIdleClick = true,
                    IdleClickPosition = pos
                };
                group.Moles.Add(idleMole);
            }
            
            // 刷新列表显示（包含序号）
            RefreshCurrentMoleList();
        }
        
        private void UpdateIdleClickLabel()
        {
            var group = GetCurrentMoleGroup();
            if (group.IdleClickPositions.Count > 0)
            {
                lblIdleClickPos.Text = $"空击: {group.IdleClickPositions.Count} 个位置";
                lblIdleClickPos.ForeColor = Color.Green;
            }
            else
            {
                lblIdleClickPos.Text = "空击: 未设置";
                lblIdleClickPos.ForeColor = Color.Gray;
            }
        }
        
        private void ChkMoleEnabled_CheckedChanged(object? sender, EventArgs e)
        {
            if (chkMoleEnabled.Checked)
            {
                var group = GetCurrentMoleGroup();
                var lstMoles = GetCurrentMoleListBox();
                
                if (group.Moles.Count == 0)
                {
                    MessageBox.Show("请先截图创建地鼠！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    chkMoleEnabled.Checked = false;
                    return;
                }
                
                // 更新地鼠启用状态
                if (lstMoles != null)
                {
                    for (int i = 0; i < group.Moles.Count && i < lstMoles.Items.Count; i++)
                    {
                        group.Moles[i].IsEnabled = lstMoles.GetItemChecked(i);
                    }
                }
                
                _moleHunter.Start(group.Moles, group.IdleClickPositions, _moleGroups);
                AppendLog($"🎯 打地鼠已启动 - 分组: {group.Name}", LogType.Success);
                if (group.IdleClickPositions.Count > 0)
                {
                    AppendLog($"💤 空击位置数量: {group.IdleClickPositions.Count}", LogType.Info);
                }
            }
            else
            {
                _moleHunter.Stop();
                AppendLog("⏸️ 打地鼠已停止", LogType.Warning);
            }
        }

        private void MoleHunter_HuntingStopped(object? sender, EventArgs e)
        {
            // 在UI线程上更新复选框状态
            if (InvokeRequired)
            {
                Invoke(new Action(() => MoleHunter_HuntingStopped(sender, e)));
                return;
            }
            
            // 取消勾选打地鼠复选框
            chkMoleEnabled.Checked = false;
        }
        

        
        private void BtnSetIdleClick_Click(object? sender, EventArgs e)
        {
            AppendLog("\n💤 请点击屏幕上的空击位置...", LogType.Info);
            AppendLog("提示: 可以设置多个位置，会循环点击", LogType.Info);
            
            // 等待用户点击
            Task.Run(async () =>
            {
                await Task.Delay(200); // 给用户200ms准备时间
                
                // 等待鼠标左键点击
                while (true)
                {
                    if ((WindowHelper.GetAsyncKeyState(WindowHelper.VK_LBUTTON) & 0x8000) != 0)
                    {
                        WindowHelper.GetCursorPos(out var pos);
                        var newPoint = new Point(pos.X, pos.Y);
                        var group = GetCurrentMoleGroup();
                        group.IdleClickPositions.Add(newPoint);
                        
                        Invoke(new Action(() =>
                        {
                            UpdateIdleClickLabel();
                            AppendLog($"✅ 空击位置 {group.IdleClickPositions.Count}: ({pos.X}, {pos.Y})", LogType.Success);
                            UpdateIdleClicksInList();
                            SaveMoles(); // 保存配置
                        }));
                        
                        break;
                    }
                    
                    await Task.Delay(50);
                }
            });
        }

        private void BtnBatchSelect_Click(object? sender, EventArgs e)
        {
            var group = GetCurrentMoleGroup();
            if (group.Moles.Count == 0)
            {
                MessageBox.Show("当前分组没有步骤", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            
            // 确保保存的位置在有效范围内
            if (_batchSelectSliderA < 1 || _batchSelectSliderA > group.Moles.Count)
                _batchSelectSliderA = 1;
            if (_batchSelectSliderB < 1 || _batchSelectSliderB > group.Moles.Count)
                _batchSelectSliderB = group.Moles.Count;
            
            // 创建批量选择对话框
            var dialog = new Form
            {
                Text = "批量启用/禁用步骤",
                Size = new Size(450, 280),
                StartPosition = FormStartPosition.Manual,
                FormBorderStyle = FormBorderStyle.FixedDialog,
                MaximizeBox = false,
                MinimizeBox = false
            };
            
            // 设置对话框位置：左边与主窗口右边对齐
            dialog.Location = new Point(this.Right, this.Top + (this.Height - dialog.Height) / 2);
            
            var lblTitle = new Label
            {
                Text = $"当前分组: {group.Name} (共 {group.Moles.Count} 个步骤)",
                Location = new Point(20, 20),
                Size = new Size(400, 20),
                Font = new Font(Font.FontFamily, 10, FontStyle.Bold),
                Parent = dialog
            };
            
            // A 滑块标签
            var lblSliderA = new Label
            {
                Text = "起始步骤 (A):",
                Location = new Point(20, 60),
                Size = new Size(100, 20),
                Parent = dialog
            };
            
            // A 滑块
            var trackBarA = new TrackBar
            {
                Location = new Point(120, 55),
                Size = new Size(280, 45),
                Minimum = 1,
                Maximum = group.Moles.Count,
                Value = _batchSelectSliderA,
                TickFrequency = Math.Max(1, group.Moles.Count / 20),
                BackColor = Color.LightBlue, // 默认蓝色（初始焦点）
                Parent = dialog
            };
            
            // A 滑块值显示
            var lblValueA = new Label
            {
                Text = _batchSelectSliderA.ToString(),
                Location = new Point(410, 60),
                Size = new Size(30, 20),
                Font = new Font(Font.FontFamily, 10, FontStyle.Bold),
                ForeColor = Color.Blue,
                Parent = dialog
            };
            
            // B 滑块标签
            var lblSliderB = new Label
            {
                Text = "结束步骤 (B):",
                Location = new Point(20, 110),
                Size = new Size(100, 20),
                Parent = dialog
            };
            
            // B 滑块
            var trackBarB = new TrackBar
            {
                Location = new Point(120, 105),
                Size = new Size(280, 45),
                Minimum = 1,
                Maximum = group.Moles.Count,
                Value = _batchSelectSliderB,
                TickFrequency = Math.Max(1, group.Moles.Count / 20),
                Parent = dialog
            };
            
            // B 滑块值显示
            var lblValueB = new Label
            {
                Text = _batchSelectSliderB.ToString(),
                Location = new Point(410, 110),
                Size = new Size(30, 20),
                Font = new Font(Font.FontFamily, 10, FontStyle.Bold),
                ForeColor = Color.Red,
                Parent = dialog
            };
            
            // A 滑块值改变事件
            trackBarA.ValueChanged += (s, ev) =>
            {
                int newValueA = trackBarA.Value;
                
                // 如果 A 尝试越过 B（A >= B），推动 B 一起移动
                if (newValueA >= _batchSelectSliderB)
                {
                    // A 推动 B，保持 B 在 A 的右边（至少相差 1）
                    _batchSelectSliderB = Math.Min(newValueA + 1, trackBarA.Maximum);
                    trackBarB.Value = _batchSelectSliderB;
                    
                    // 如果 B 已经到达最大值，限制 A 的位置
                    if (_batchSelectSliderB == trackBarA.Maximum)
                    {
                        newValueA = _batchSelectSliderB - 1;
                        trackBarA.Value = newValueA;
                    }
                }
                
                _batchSelectSliderA = newValueA;
                lblValueA.Text = _batchSelectSliderA.ToString();
                lblValueB.Text = _batchSelectSliderB.ToString();
            };
            
            // B 滑块值改变事件
            trackBarB.ValueChanged += (s, ev) =>
            {
                int newValueB = trackBarB.Value;
                
                // 如果 B 尝试越过 A（B <= A），推动 A 一起移动
                if (newValueB <= _batchSelectSliderA)
                {
                    // B 推动 A，保持 A 在 B 的左边（至少相差 1）
                    _batchSelectSliderA = Math.Max(newValueB - 1, trackBarB.Minimum);
                    trackBarA.Value = _batchSelectSliderA;
                    
                    // 如果 A 已经到达最小值，限制 B 的位置
                    if (_batchSelectSliderA == trackBarB.Minimum)
                    {
                        newValueB = _batchSelectSliderA + 1;
                        trackBarB.Value = newValueB;
                    }
                }
                
                _batchSelectSliderB = newValueB;
                lblValueA.Text = _batchSelectSliderA.ToString();
                lblValueB.Text = _batchSelectSliderB.ToString();
            };
            
            // A 滑块获得焦点事件
            trackBarA.Enter += (s, ev) =>
            {
                trackBarA.BackColor = Color.LightBlue;
                trackBarB.BackColor = SystemColors.Control; // 恢复默认色
            };
            
            // B 滑块获得焦点事件
            trackBarB.Enter += (s, ev) =>
            {
                trackBarB.BackColor = Color.LightBlue;
                trackBarA.BackColor = SystemColors.Control; // 恢复默认色
            };
            
            // A 滑块键盘事件
            trackBarA.KeyDown += (s, ev) =>
            {
                if (ev.KeyCode == Keys.Down)
                {
                    // 按下键，切换到 B 滑块
                    trackBarB.Focus();
                    ev.Handled = true;
                }
            };
            
            // B 滑块键盘事件
            trackBarB.KeyDown += (s, ev) =>
            {
                if (ev.KeyCode == Keys.Up)
                {
                    // 按上键，切换到 A 滑块
                    trackBarA.Focus();
                    ev.Handled = true;
                }
            };
            
            // 提示标签
            var lblHint = new Label
            {
                Text = "拖动滑块或使用左右键调整位置，上下键切换滑块",
                Location = new Point(20, 165),
                Size = new Size(400, 20),
                ForeColor = Color.Gray,
                Parent = dialog
            };
            
            // 全部启用按钮
            var btnEnableAll = new Button
            {
                Text = "启用全部",
                Location = new Point(70, 200),
                Size = new Size(150, 35),
                Parent = dialog
            };
            
            // 全部禁用按钮
            var btnDisableAll = new Button
            {
                Text = "禁用 A-B 之间的步骤",
                Location = new Point(230, 200),
                Size = new Size(150, 35),
                Parent = dialog
            };
            
            // 启用全部按钮点击事件
            btnEnableAll.Click += (s, ev) =>
            {
                int count = 0;
                
                // 启用所有步骤
                for (int i = 0; i < group.Moles.Count; i++)
                {
                    group.Moles[i].IsEnabled = true;
                    count++;
                }
                
                SaveMoles();
                RefreshCurrentMoleList();
                AppendLog($"✅ 已启用全部步骤，共 {count} 个", LogType.Success);
                dialog.Close();
            };
            
            // 禁用按钮点击事件
            btnDisableAll.Click += (s, ev) =>
            {
                int start = Math.Min(_batchSelectSliderA, _batchSelectSliderB) - 1; // 转换为索引
                int end = Math.Max(_batchSelectSliderA, _batchSelectSliderB) - 1;
                int count = 0;
                
                for (int i = start; i <= end && i < group.Moles.Count; i++)
                {
                    group.Moles[i].IsEnabled = false;
                    count++;
                }
                
                SaveMoles();
                RefreshCurrentMoleList();
                AppendLog($"✅ 已禁用步骤 {start + 1} 到 {end + 1}，共 {count} 个步骤", LogType.Success);
                dialog.Close();
            };
            
            dialog.ShowDialog();
            // 对话框关闭后，位置已经保存在 _batchSelectSliderA 和 _batchSelectSliderB 中
        }
        
        private void BtnAddJump_Click(object? sender, EventArgs e)
        {
            // 获取所有分组名称，除了当前分组
            var currentGroup = GetCurrentMoleGroup();
            var otherGroups = _moleGroups
                .Where(g => g.Name != currentGroup.Name)
                .ToList();

            if (otherGroups.Count == 0)
            {
                MessageBox.Show("没有其他分组可以跳转到", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            // 创建选择框
            var form = new Form
            {
                Text = "选择跳转目标",
                Size = new Size(350, 280),
                StartPosition = FormStartPosition.Manual,
                FormBorderStyle = FormBorderStyle.FixedDialog,
                MaximizeBox = false,
                MinimizeBox = false
            };
            
            // 设置对话框位置：左边与主窗口右边对齐
            form.Location = new Point(this.Right, this.Top + (this.Height - form.Height) / 2);

            var label1 = new Label
            {
                Text = "选择要跳转到的分组:",
                Location = new Point(20, 20),
                Size = new Size(310, 20),
                Parent = form
            };

            var comboGroup = new ComboBox
            {
                Location = new Point(20, 45),
                Size = new Size(310, 25),
                DropDownStyle = ComboBoxStyle.DropDownList,
                Parent = form
            };

            foreach (var group in otherGroups)
            {
                comboGroup.Items.Add(group.Name);
            }

            if (comboGroup.Items.Count > 0)
                comboGroup.SelectedIndex = 0;

            var label2 = new Label
            {
                Text = "选择目标分组中的步骤 (可选):",
                Location = new Point(20, 85),
                Size = new Size(310, 20),
                Parent = form
            };

            var comboStep = new ComboBox
            {
                Location = new Point(20, 110),
                Size = new Size(310, 25),
                DropDownStyle = ComboBoxStyle.DropDownList,
                Parent = form
            };

            // 当分组选择改变时，更新步骤列表
            comboGroup.SelectedIndexChanged += (s, e) =>
            {
                comboStep.Items.Clear();
                comboStep.Items.Add("(从头开始)");
                
                if (comboGroup.SelectedIndex >= 0 && comboGroup.SelectedIndex < otherGroups.Count)
                {
                    var selectedGroup = otherGroups[comboGroup.SelectedIndex];
                    for (int i = 0; i < selectedGroup.Moles.Count; i++)
                    {
                        var mole = selectedGroup.Moles[i];
                        var displayName = mole.IsIdleClick && mole.IdleClickPosition.HasValue
                            ? $"{i + 1}. 💤 {mole.Name}"
                            : mole.IsJump
                            ? $"{i + 1}. 🔗 {mole.Name}"
                            : $"{i + 1}. {mole.Name}";
                        comboStep.Items.Add(displayName);
                    }
                }
                
                comboStep.SelectedIndex = 0;
            };

            // 初始化步骤列表
            if (comboGroup.SelectedIndex >= 0)
            {
                comboGroup_SelectedIndexChanged(null, EventArgs.Empty);
            }

            var hintLabel = new Label
            {
                Text = "提示: 不选择步骤则从分组开始执行；选择步骤则从该步骤开始执行",
                Location = new Point(20, 145),
                Size = new Size(310, 40),
                ForeColor = Color.Gray,
                AutoSize = false,
                Parent = form
            };

            var btnOk = new Button
            {
                Text = "确定",
                Location = new Point(150, 200),
                Size = new Size(80, 30),
                DialogResult = DialogResult.OK,
                Parent = form
            };

            var btnCancel = new Button
            {
                Text = "取消",
                Location = new Point(240, 200),
                Size = new Size(80, 30),
                DialogResult = DialogResult.Cancel,
                Parent = form
            };

            form.AcceptButton = btnOk;
            form.CancelButton = btnCancel;

            // 处理分组选择变化的事件
            void comboGroup_SelectedIndexChanged(object? s, EventArgs e)
            {
                comboStep.Items.Clear();
                comboStep.Items.Add("(从头开始)");
                
                if (comboGroup.SelectedIndex >= 0 && comboGroup.SelectedIndex < otherGroups.Count)
                {
                    var selectedGroup = otherGroups[comboGroup.SelectedIndex];
                    for (int i = 0; i < selectedGroup.Moles.Count; i++)
                    {
                        var mole = selectedGroup.Moles[i];
                        var displayName = mole.IsIdleClick && mole.IdleClickPosition.HasValue
                            ? $"{i + 1}. 💤 {mole.Name}"
                            : mole.IsJump
                            ? $"{i + 1}. 🔗 {mole.Name}"
                            : $"{i + 1}. {mole.Name}";
                        comboStep.Items.Add(displayName);
                    }
                }
                
                comboStep.SelectedIndex = 0;
            }

            if (form.ShowDialog() == DialogResult.OK && comboGroup.SelectedIndex >= 0)
            {
                var targetGroupName = comboGroup.SelectedItem.ToString();
                var stepIndex = comboStep.SelectedIndex - 1; // -1 表示从头开始
                
                // 创建跳转步骤
                var jumpMole = new MoleItem
                {
                    Name = stepIndex < 0 
                        ? $"🔗 跳转到 {targetGroupName}" 
                        : $"🔗 跳转到 {targetGroupName} (步骤 {stepIndex + 1})",
                    IsJump = true,
                    JumpTargetGroup = targetGroupName,
                    JumpTargetStep = stepIndex,
                    IsEnabled = true
                };

                currentGroup.Moles.Add(jumpMole);
                SaveMoles();
                
                // 更新列表显示
                var lstMoles = GetCurrentMoleListBox();
                if (lstMoles != null)
                {
                    lstMoles.Items.Add(jumpMole.Name, true);
                }

                var stepInfo = stepIndex < 0 ? "从头开始" : $"从步骤 {stepIndex + 1} 开始";
                AppendLog($"✅ 已添加跳转步骤: 跳转到 {targetGroupName} ({stepInfo})", LogType.Success);
            }
        }

        private void BtnCaptureMole_Click(object? sender, EventArgs e)
        {
            // 最小化窗口
            WindowState = FormWindowState.Minimized;
            Thread.Sleep(500); // 等待窗口最小化
            
            // 截图
            var screenshot = CaptureScreen();
            
            // 恢复窗口
            WindowState = FormWindowState.Normal;
            
            // 显示截图选择对话框
            var dialog = new Form
            {
                Text = "选择地鼠区域",
                Size = new Size(Screen.PrimaryScreen.Bounds.Width, Screen.PrimaryScreen.Bounds.Height),
                StartPosition = FormStartPosition.Manual,
                Location = new Point(0, 0),
                FormBorderStyle = FormBorderStyle.None,
                WindowState = FormWindowState.Maximized,
                BackgroundImage = screenshot,
                BackgroundImageLayout = ImageLayout.Stretch
            };
            
            Point? startPoint = null;
            Rectangle? selection = null;
            
            dialog.MouseDown += (s, me) =>
            {
                if (me.Button == MouseButtons.Left)
                {
                    startPoint = me.Location;
                }
            };
            
            dialog.MouseMove += (s, me) =>
            {
                if (startPoint.HasValue)
                {
                    var rect = GetRectangle(startPoint.Value, me.Location);
                    selection = rect;
                    dialog.Invalidate();
                }
            };
            
            dialog.MouseUp += (s, me) =>
            {
                if (me.Button == MouseButtons.Left && selection.HasValue)
                {
                    dialog.DialogResult = DialogResult.OK;
                    dialog.Close();
                }
            };
            
            dialog.Paint += (s, pe) =>
            {
                if (selection.HasValue)
                {
                    using (var pen = new Pen(Color.Red, 2))
                    {
                        pe.Graphics.DrawRectangle(pen, selection.Value);
                    }
                }
            };
            
            dialog.KeyDown += (s, ke) =>
            {
                if (ke.KeyCode == Keys.Escape)
                {
                    dialog.DialogResult = DialogResult.Cancel;
                    dialog.Close();
                }
            };
            
            if (dialog.ShowDialog() == DialogResult.OK && selection.HasValue)
            {
                // 裁剪图像
                var croppedImage = CropImage(screenshot, selection.Value);
                
                // 保存图像
                var fileName = $"mole_{DateTime.Now:yyyyMMdd_HHmmss}.png";
                var filePath = Path.Combine(_molesDirectory, fileName);
                croppedImage.Save(filePath, System.Drawing.Imaging.ImageFormat.Png);
                
                // 创建新的地鼠项
                var newMole = new MoleItem
                {
                    Name = Path.GetFileNameWithoutExtension(fileName),
                    ImagePath = filePath,
                    IsEnabled = true,
                    SimilarityThreshold = 0.85,
                    WaitUntilAppear = true // 默认选中"持续等待直到出现"
                };
                
                // 添加到当前分组
                var group = GetCurrentMoleGroup();
                group.Moles.Add(newMole);
                
                // 保存配置
                SaveMoles();
                
                // 刷新当前列表显示（包含序号）
                RefreshCurrentMoleList();
                
                AppendLog($"✅ 已创建地鼠: {fileName} (分组: {group.Name})", LogType.Success);
            }
            
            screenshot.Dispose();
        }
        
        private Bitmap CaptureScreen()
        {
            var bounds = Screen.PrimaryScreen.Bounds;
            var bitmap = new Bitmap(bounds.Width, bounds.Height);
            
            using (var g = Graphics.FromImage(bitmap))
            {
                g.CopyFromScreen(Point.Empty, Point.Empty, bounds.Size);
            }
            
            return bitmap;
        }
        
        private Rectangle GetRectangle(Point p1, Point p2)
        {
            return new Rectangle(
                Math.Min(p1.X, p2.X),
                Math.Min(p1.Y, p2.Y),
                Math.Abs(p1.X - p2.X),
                Math.Abs(p1.Y - p2.Y)
            );
        }
        
        private Bitmap CropImage(Bitmap source, Rectangle cropArea)
        {
            var cropped = new Bitmap(cropArea.Width, cropArea.Height);
            
            using (var g = Graphics.FromImage(cropped))
            {
                g.DrawImage(source, 
                    new Rectangle(0, 0, cropArea.Width, cropArea.Height),
                    cropArea,
                    GraphicsUnit.Pixel);
            }
            
            return cropped;
        }
        
        // 预览窗口相关字段
        private Form? _previewForm = null;
        private PictureBox? _previewPictureBox = null;
        private Label? _previewStepLabel = null;
        private int _lastPreviewIndex = -1;
        private int _hoveredMoleIndex = -1;
        private CheckedListBox? _lastHoveredListBox = null;
        
        private void LstMoles_MouseLeave(object? sender, EventArgs e)
        {
            if (sender is CheckedListBox lstMoles)
            {
                // 鼠标离开列表时，清除悬浮状态和预览
                HidePreview();
                UpdateHoveredItem(lstMoles, -1);
            }
        }
        
        private void LstMoles_MouseMove(object? sender, MouseEventArgs e)
        {
            if (sender is CheckedListBox lstMoles)
            {
                var group = GetCurrentMoleGroup();
                var index = lstMoles.IndexFromPoint(e.Location);
                
                // 如果鼠标移出列表项或索引无效，隐藏预览
                if (index < 0 || index >= group.Moles.Count)
                {
                    HidePreview();
                    UpdateHoveredItem(lstMoles, -1);
                    return;
                }
                
                // 更新悬浮项（触发重绘）
                UpdateHoveredItem(lstMoles, index);
                
                // 如果是同一个项，不需要重新显示预览
                if (index == _lastPreviewIndex)
                    return;
                
                _lastPreviewIndex = index;
                var mole = group.Moles[index];
                
                // 只为截图地鼠显示预览
                if (!mole.IsIdleClick && !mole.IsJump && !string.IsNullOrEmpty(mole.ImagePath) && File.Exists(mole.ImagePath))
                {
                    ShowPreview(mole.ImagePath, lstMoles);
                }
                else
                {
                    HidePreview();
                }
            }
        }
        
        private void UpdateHoveredItem(CheckedListBox lstMoles, int newIndex)
        {
            if (_hoveredMoleIndex != newIndex || _lastHoveredListBox != lstMoles)
            {
                // 重绘旧的悬浮项
                if (_lastHoveredListBox != null && _hoveredMoleIndex >= 0 && _hoveredMoleIndex < _lastHoveredListBox.Items.Count)
                {
                    var oldRect = _lastHoveredListBox.GetItemRectangle(_hoveredMoleIndex);
                    _lastHoveredListBox.Invalidate(oldRect);
                    _lastHoveredListBox.Update(); // 强制立即重绘
                }
                
                _hoveredMoleIndex = newIndex;
                _lastHoveredListBox = lstMoles;
                
                // 重绘新的悬浮项
                if (_hoveredMoleIndex >= 0 && _hoveredMoleIndex < lstMoles.Items.Count)
                {
                    var newRect = lstMoles.GetItemRectangle(_hoveredMoleIndex);
                    lstMoles.Invalidate(newRect);
                    lstMoles.Update(); // 强制立即重绘
                }
            }
        }
        
        private void ShowPreview(string imagePath, Control relativeControl)
        {
            try
            {
                // 创建预览窗口（如果不存在）
                if (_previewForm == null)
                {
                    _previewForm = new Form
                    {
                        FormBorderStyle = FormBorderStyle.None,
                        StartPosition = FormStartPosition.Manual,
                        ShowInTaskbar = false,
                        TopMost = true,
                        BackColor = Color.White,
                        Padding = new Padding(2)
                    };
                    
                    // 序号标签（显示在顶部）
                    _previewStepLabel = new Label
                    {
                        Dock = DockStyle.Top,
                        Height = 25,
                        TextAlign = System.Drawing.ContentAlignment.MiddleCenter,
                        Font = new Font("Microsoft YaHei UI", 10, FontStyle.Bold),
                        BackColor = Color.FromArgb(0, 120, 215), // Windows 蓝色
                        ForeColor = Color.White,
                        Parent = _previewForm
                    };
                    
                    _previewPictureBox = new PictureBox
                    {
                        SizeMode = PictureBoxSizeMode.Zoom,
                        Dock = DockStyle.Fill,
                        Parent = _previewForm
                    };
                    
                    // 当鼠标离开预览窗口时隐藏
                    _previewForm.MouseLeave += (s, e) =>
                    {
                        var clientPoint = _previewForm.PointToClient(Cursor.Position);
                        if (!_previewForm.ClientRectangle.Contains(clientPoint))
                        {
                            HidePreview();
                        }
                    };
                }
                
                // 加载图片
                if (_previewPictureBox?.Image != null)
                {
                    var oldImage = _previewPictureBox.Image;
                    _previewPictureBox.Image = null;
                    oldImage.Dispose();
                }
                
                var image = Image.FromFile(imagePath);
                _previewPictureBox!.Image = image;
                
                // 更新序号标签文本
                if (_previewStepLabel != null && _hoveredMoleIndex >= 0)
                {
                    _previewStepLabel.Text = $"步骤 {_hoveredMoleIndex + 1}";
                }
                
                // 计算预览窗口大小（最大 300x300，加上标签高度）
                int maxSize = 300;
                double scale = Math.Min((double)maxSize / image.Width, (double)maxSize / image.Height);
                if (scale > 1) scale = 1; // 不放大
                
                int previewWidth = (int)(image.Width * scale) + 4; // +4 for padding
                int previewHeight = (int)(image.Height * scale) + 4 + 25; // +25 for label height
                
                _previewForm.Size = new Size(previewWidth, previewHeight);
                
                // 计算预览窗口位置（显示在列表右侧）
                var screenPoint = relativeControl.PointToScreen(new Point(relativeControl.Width + 10, Cursor.Position.Y - relativeControl.PointToScreen(Point.Empty).Y));
                
                // 确保预览窗口不超出屏幕
                var screen = Screen.FromControl(relativeControl);
                if (screenPoint.X + previewWidth > screen.WorkingArea.Right)
                {
                    screenPoint.X = relativeControl.PointToScreen(Point.Empty).X - previewWidth - 10;
                }
                if (screenPoint.Y + previewHeight > screen.WorkingArea.Bottom)
                {
                    screenPoint.Y = screen.WorkingArea.Bottom - previewHeight;
                }
                
                _previewForm.Location = screenPoint;
                _previewForm.Show();
            }
            catch
            {
                HidePreview();
            }
        }
        
        private void HidePreview()
        {
            _lastPreviewIndex = -1;
            
            if (_previewForm != null)
            {
                _previewForm.Hide();
                
                if (_previewPictureBox?.Image != null)
                {
                    var oldImage = _previewPictureBox.Image;
                    _previewPictureBox.Image = null;
                    oldImage.Dispose();
                }
            }
            
            // 注意：不清除悬浮状态，让步骤保持红色高亮
            // 悬浮状态只在鼠标移出列表时才清除
        }
        
        private void LstMoles_DrawItem(object? sender, DrawItemEventArgs e)
        {
            if (sender is CheckedListBox lstMoles && e.Index >= 0 && e.Index < lstMoles.Items.Count)
            {
                // 判断是否是悬浮项
                bool isHovered = (e.Index == _hoveredMoleIndex && lstMoles == _lastHoveredListBox);
                
                // 手动绘制背景（使用主题颜色）
                Color backColor;
                if ((e.State & DrawItemState.Selected) == DrawItemState.Selected)
                {
                    // 选中项使用高亮背景
                    backColor = SystemColors.Highlight;
                }
                else if (isHovered)
                {
                    // 悬浮项使用浅黄色高亮背景
                    var effectiveTheme = _themeManager.GetEffectiveTheme();
                    if (effectiveTheme == ThemeMode.Dark)
                    {
                        // 深色主题：使用深橙色
                        backColor = Color.FromArgb(80, 60, 30);
                    }
                    else
                    {
                        // 浅色主题：使用浅黄色
                        backColor = Color.FromArgb(255, 255, 200);
                    }
                }
                else
                {
                    // 未选中项使用控件的背景色（已被主题管理器设置）
                    backColor = lstMoles.BackColor;
                }
                
                using (var backBrush = new SolidBrush(backColor))
                {
                    e.Graphics.FillRectangle(backBrush, e.Bounds);
                }
                
                // 绘制复选框
                var checkBoxRect = new Rectangle(e.Bounds.Left + 2, e.Bounds.Top + 2, 16, 16);
                var checkState = lstMoles.GetItemChecked(e.Index) ? System.Windows.Forms.VisualStyles.CheckBoxState.CheckedNormal : System.Windows.Forms.VisualStyles.CheckBoxState.UncheckedNormal;
                CheckBoxRenderer.DrawCheckBox(e.Graphics, checkBoxRect.Location, checkState);
                
                // 获取文本内容
                string fullText = lstMoles.Items[e.Index].ToString() ?? "";
                
                // 分离序号和内容（序号格式：数字 + "."）
                string numberPart = "";
                string contentPart = fullText;
                int dotIndex = fullText.IndexOf('.');
                if (dotIndex > 0)
                {
                    numberPart = fullText.Substring(0, dotIndex + 1); // 包含点号
                    contentPart = fullText.Substring(dotIndex + 1); // 点号后的内容
                }
                
                // 确定文本颜色：悬浮时为红色，选中时为高亮文本色，否则使用控件前景色
                Color textColor;
                if (isHovered)
                {
                    textColor = Color.Red; // 悬浮时显示红色
                }
                else if ((e.State & DrawItemState.Selected) == DrawItemState.Selected)
                {
                    textColor = SystemColors.HighlightText;
                }
                else
                {
                    // 使用控件的前景色（已被主题管理器设置）
                    textColor = lstMoles.ForeColor;
                }
                
                // 绘制序号（悬浮时使用更大的字体）
                int xOffset = e.Bounds.Left + 22;
                if (!string.IsNullOrEmpty(numberPart))
                {
                    Font numberFont = isHovered 
                        ? new Font(e.Font.FontFamily, e.Font.Size + 2, FontStyle.Bold) 
                        : e.Font;
                    
                    var numberSize = TextRenderer.MeasureText(e.Graphics, numberPart, numberFont);
                    var numberRect = new Rectangle(
                        xOffset,
                        e.Bounds.Top,
                        numberSize.Width,
                        e.Bounds.Height
                    );
                    
                    TextRenderer.DrawText(
                        e.Graphics,
                        numberPart,
                        numberFont,
                        numberRect,
                        textColor,
                        TextFormatFlags.Left | TextFormatFlags.VerticalCenter
                    );
                    
                    if (isHovered)
                    {
                        numberFont.Dispose();
                    }
                    
                    xOffset += numberSize.Width;
                }
                
                // 绘制内容部分
                var contentRect = new Rectangle(
                    xOffset,
                    e.Bounds.Top,
                    e.Bounds.Width - (xOffset - e.Bounds.Left),
                    e.Bounds.Height
                );
                
                TextRenderer.DrawText(
                    e.Graphics,
                    contentPart,
                    e.Font,
                    contentRect,
                    textColor,
                    TextFormatFlags.Left | TextFormatFlags.VerticalCenter
                );
                
                // 绘制悬浮边框
                if (isHovered)
                {
                    using (var pen = new Pen(Color.OrangeRed, 2))
                    {
                        var borderRect = new Rectangle(
                            e.Bounds.Left + 1,
                            e.Bounds.Top + 1,
                            e.Bounds.Width - 2,
                            e.Bounds.Height - 2
                        );
                        e.Graphics.DrawRectangle(pen, borderRect);
                    }
                }
                
                // 绘制焦点框
                e.DrawFocusRectangle();
            }
        }
        
        private void LstMoles_KeyDown(object? sender, KeyEventArgs e)
        {
            if (sender is CheckedListBox lstMoles)
            {
                var group = GetCurrentMoleGroup();
                
                // 获取当前选中的索引
                if (lstMoles.SelectedIndex < 0 || lstMoles.SelectedIndex >= group.Moles.Count)
                    return;
                
                int currentIndex = lstMoles.SelectedIndex;
                int newIndex = -1;
                
                // 处理上下键
                if (e.KeyCode == Keys.Up && currentIndex > 0)
                {
                    // 向上移动
                    newIndex = currentIndex - 1;
                    e.Handled = true;
                    e.SuppressKeyPress = true;
                }
                else if (e.KeyCode == Keys.Down && currentIndex < group.Moles.Count - 1)
                {
                    // 向下移动
                    newIndex = currentIndex + 1;
                    e.Handled = true;
                    e.SuppressKeyPress = true;
                }
                
                // 如果需要移动
                if (newIndex >= 0)
                {
                    // 交换地鼠在列表中的位置
                    var mole = group.Moles[currentIndex];
                    group.Moles.RemoveAt(currentIndex);
                    group.Moles.Insert(newIndex, mole);
                    
                    // 保存配置
                    SaveMoles();
                    
                    // 刷新列表显示
                    RefreshCurrentMoleList();
                    
                    // 重新选中移动后的项
                    lstMoles.SelectedIndex = newIndex;
                    
                    AppendLog($"✅ 已移动步骤: {mole.Name} (从位置 {currentIndex + 1} 到 {newIndex + 1})", LogType.Success);
                }
            }
        }
        
        private void RefreshCurrentMoleList()
        {
            var lstMoles = GetCurrentMoleListBox();
            if (lstMoles == null)
                return;
            
            var group = GetCurrentMoleGroup();
            
            // 保存当前的选中索引
            int selectedIndex = lstMoles.SelectedIndex;
            
            // 清空并重新加载列表
            lstMoles.Items.Clear();
            
            for (int i = 0; i < group.Moles.Count; i++)
            {
                var mole = group.Moles[i];
                string displayText;
                
                if (mole.IsIdleClick && mole.IdleClickPosition.HasValue)
                {
                    displayText = $"{i + 1}. 💤 {mole.Name}: ({mole.IdleClickPosition.Value.X}, {mole.IdleClickPosition.Value.Y})";
                }
                else if (mole.IsJump)
                {
                    displayText = $"{i + 1}. 🔗 {mole.Name}";
                }
                else
                {
                    displayText = $"{i + 1}. {mole.Name}";
                }
                
                lstMoles.Items.Add(displayText, mole.IsEnabled);
            }
        }
        
        private void LstMoles_MouseDown(object? sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right && sender is CheckedListBox lstMoles)
            {
                var group = GetCurrentMoleGroup();
                if (group == null) return;
                
                var index = lstMoles.IndexFromPoint(e.Location);
                
                if (index >= 0 && index < group.Moles.Count)
                {
                    // 右键点击了某个步骤，关闭当前编辑窗口并打开新的
                    CloseCurrentEditDialog();
                    
                    var mole = group.Moles[index];
                    
                    // 如果是跳转步骤，显示编辑对话框
                    if (mole.IsJump)
                    {
                        ShowJumpStepEditDialog(mole, index);
                        return;
                    }
                    
                    // 如果是空击地鼠，显示自定义对话框
                    if (mole.IsIdleClick)
                    {
                        ShowIdleClickEditDialog(mole, index);
                        return;
                    }
                    
                    // 创建自定义确认对话框，显示预览图（非模态）
                    ShowMoleDeleteConfirmDialog(mole, index);
                }
                else
                {
                    // 右键点击了空白处，关闭当前编辑窗口
                    CloseCurrentEditDialog();
                }
            }
        }
        
        private void CloseCurrentEditDialog()
        {
            try
            {
                if (_currentEditDialog != null && !_currentEditDialog.IsDisposed)
                {
                    _currentEditDialog.Close();
                    _currentEditDialog.Dispose();
                }
            }
            catch (Exception ex)
            {
                // 忽略关闭窗口时的异常
                System.Diagnostics.Debug.WriteLine($"关闭编辑窗口异常: {ex.Message}");
            }
            finally
            {
                _currentEditDialog = null;
            }
        }
        
        private void ShowIdleClickEditDialog(MoleItem idleMole, int moleIndex)
        {
            var currentGroup = GetCurrentMoleGroup();
            
            // 创建编辑对话框
            var form = new Form
            {
                Text = "空击步骤设置",
                Size = new Size(400, 250),
                StartPosition = FormStartPosition.Manual,
                FormBorderStyle = FormBorderStyle.FixedDialog,
                MaximizeBox = false,
                MinimizeBox = false
            };
            
            // 设置对话框位置：左边与主窗口右边对齐
            form.Location = new Point(this.Right, this.Top + (this.Height - form.Height) / 2);
            
            var lblInfo = new Label
            {
                Text = $"空击位置: {idleMole.Name}",
                Location = new Point(20, 20),
                Size = new Size(350, 20),
                Font = new Font(Font.FontFamily, 10, FontStyle.Bold),
                Parent = form
            };
            
            var lblPosition = new Label
            {
                Text = $"坐标: ({idleMole.IdleClickPosition?.X}, {idleMole.IdleClickPosition?.Y})",
                Location = new Point(20, 50),
                Size = new Size(350, 20),
                ForeColor = Color.Gray,
                Parent = form
            };
            
            // 停止打地鼠复选框
            var chkStopHunting = new CheckBox
            {
                Text = "执行到此步骤时停止打地鼠",
                Location = new Point(20, 90),
                Size = new Size(350, 25),
                Checked = idleMole.StopHunting,
                Parent = form
            };
            
            var lblHint = new Label
            {
                Text = "选中后，执行到此步骤时会自动停止打地鼠，不执行点击",
                Location = new Point(40, 115),
                Size = new Size(330, 40),
                ForeColor = Color.Gray,
                Font = new Font(Font.FontFamily, 8),
                Parent = form
            };
            
            var btnSave = new Button
            {
                Text = "保存",
                Location = new Point(100, 170),
                Size = new Size(80, 30),
                Parent = form
            };
            
            var btnDelete = new Button
            {
                Text = "删除",
                Location = new Point(190, 170),
                Size = new Size(80, 30),
                Parent = form
            };
            
            var btnCancel = new Button
            {
                Text = "取消",
                Location = new Point(280, 170),
                Size = new Size(80, 30),
                DialogResult = DialogResult.Cancel,
                Parent = form
            };
            
            // 保存按钮点击事件
            btnSave.Click += (s, e) =>
            {
                idleMole.StopHunting = chkStopHunting.Checked;
                SaveMoles();
                AppendLog($"✅ 已更新空击步骤设置: {idleMole.Name}", LogType.Success);
                form.Close();
            };
            
            // 删除按钮点击事件
            btnDelete.Click += (s, e) =>
            {
                var result = MessageBox.Show(
                    $"确定要删除空击位置 \"{idleMole.Name}\" 吗？",
                    "确认删除",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Question);
                
                if (result == DialogResult.Yes)
                {
                    // 从空击位置列表中移除
                    if (idleMole.IdleClickPosition.HasValue)
                    {
                        var posToRemove = idleMole.IdleClickPosition.Value;
                        // 查找并删除匹配的位置
                        for (int i = currentGroup.IdleClickPositions.Count - 1; i >= 0; i--)
                        {
                            if (currentGroup.IdleClickPositions[i].X == posToRemove.X && 
                                currentGroup.IdleClickPositions[i].Y == posToRemove.Y)
                            {
                                currentGroup.IdleClickPositions.RemoveAt(i);
                                break;
                            }
                        }
                    }
                    
                    AppendLog($"✅ 已删除空击位置: {idleMole.Name}", LogType.Success);
                    UpdateIdleClicksInList();
                    UpdateIdleClickLabel();
                    SaveMoles();
                    form.Close();
                }
            };
            
            // 保存当前编辑窗口引用
            _currentEditDialog = form;
            
            // 窗口关闭时清除引用
            form.FormClosed += (s, e) =>
            {
                if (_currentEditDialog == form)
                {
                    _currentEditDialog = null;
                }
            };
            
            form.Show();
        }
        
        private void ShowJumpStepEditDialog(MoleItem jumpMole, int moleIndex)
        {
            var currentGroup = GetCurrentMoleGroup();
            var otherGroups = _moleGroups
                .Where(g => g.Name != currentGroup.Name)
                .ToList();

            if (otherGroups.Count == 0)
            {
                MessageBox.Show("没有其他分组可以跳转到", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            // 创建编辑对话框（加宽150用于预览）
            var form = new Form
            {
                Text = "编辑跳转步骤",
                Size = new Size(500, 380),
                StartPosition = FormStartPosition.Manual,
                FormBorderStyle = FormBorderStyle.FixedDialog,
                MaximizeBox = false,
                MinimizeBox = false
            };
            
            // 设置对话框位置：左边与主窗口右边对齐
            form.Location = new Point(this.Right, this.Top + (this.Height - form.Height) / 2);

            var label1 = new Label
            {
                Text = "选择要跳转到的分组:",
                Location = new Point(20, 20),
                Size = new Size(310, 20),
                Parent = form
            };

            var comboGroup = new ComboBox
            {
                Location = new Point(20, 45),
                Size = new Size(310, 25),
                DropDownStyle = ComboBoxStyle.DropDownList,
                Parent = form
            };

            foreach (var group in otherGroups)
            {
                comboGroup.Items.Add(group.Name);
            }

            // 设置当前选中的分组
            int currentGroupIndex = otherGroups.FindIndex(g => g.Name == jumpMole.JumpTargetGroup);
            if (currentGroupIndex >= 0)
                comboGroup.SelectedIndex = currentGroupIndex;
            else if (comboGroup.Items.Count > 0)
                comboGroup.SelectedIndex = 0;

            var label2 = new Label
            {
                Text = "选择目标分组中的步骤 (可选):",
                Location = new Point(20, 85),
                Size = new Size(310, 20),
                Parent = form
            };

            var comboStep = new ComboBox
            {
                Location = new Point(20, 110),
                Size = new Size(310, 25),
                DropDownStyle = ComboBoxStyle.DropDownList,
                Parent = form
            };

            // 预览区域
            var picPreview = new PictureBox
            {
                Location = new Point(350, 20),
                Size = new Size(130, 130),
                BorderStyle = BorderStyle.FixedSingle,
                SizeMode = PictureBoxSizeMode.Zoom,
                BackColor = Color.LightGray,
                Parent = form
            };

            var lblPreviewTitle = new Label
            {
                Text = "截图预览:",
                Location = new Point(350, 0),
                Size = new Size(130, 15),
                Font = new Font(Font.FontFamily, 9, FontStyle.Bold),
                Parent = form
            };

            // 当分组选择改变时，更新步骤列表
            comboGroup.SelectedIndexChanged += (s, e) =>
            {
                comboStep.Items.Clear();
                comboStep.Items.Add("(从头开始)");
                
                if (comboGroup.SelectedIndex >= 0 && comboGroup.SelectedIndex < otherGroups.Count)
                {
                    var selectedGroup = otherGroups[comboGroup.SelectedIndex];
                    for (int i = 0; i < selectedGroup.Moles.Count; i++)
                    {
                        var mole = selectedGroup.Moles[i];
                        var displayName = mole.IsIdleClick && mole.IdleClickPosition.HasValue
                            ? $"{i + 1}. 💤 {mole.Name}"
                            : mole.IsJump
                            ? $"{i + 1}. 🔗 {mole.Name}"
                            : $"{i + 1}. {mole.Name}";
                        comboStep.Items.Add(displayName);
                    }
                }
                
                // 恢复之前的步骤选择
                if (comboGroup.SelectedIndex >= 0 && comboGroup.SelectedIndex == currentGroupIndex)
                {
                    int stepIndex = jumpMole.JumpTargetStep + 1; // +1 因为第一项是"从头开始"
                    if (stepIndex >= 0 && stepIndex < comboStep.Items.Count)
                        comboStep.SelectedIndex = stepIndex;
                    else
                        comboStep.SelectedIndex = 0;
                }
                else
                {
                    comboStep.SelectedIndex = 0;
                }
            };

            // 当步骤选择改变时，更新预览
            comboStep.SelectedIndexChanged += (s, e) =>
            {
                // 清空预览
                if (picPreview.Image != null)
                {
                    var oldImage = picPreview.Image;
                    picPreview.Image = null;
                    oldImage.Dispose();
                }

                // 如果选择了具体步骤（不是"从头开始"），显示预览
                if (comboStep.SelectedIndex > 0 && comboGroup.SelectedIndex >= 0 && comboGroup.SelectedIndex < otherGroups.Count)
                {
                    var selectedGroup = otherGroups[comboGroup.SelectedIndex];
                    int stepIndex = comboStep.SelectedIndex - 1; // -1 因为第一项是"从头开始"
                    
                    if (stepIndex >= 0 && stepIndex < selectedGroup.Moles.Count)
                    {
                        var mole = selectedGroup.Moles[stepIndex];
                        
                        // 如果是截图步骤，显示预览
                        if (!mole.IsIdleClick && !mole.IsJump && !string.IsNullOrEmpty(mole.ImagePath) && File.Exists(mole.ImagePath))
                        {
                            try
                            {
                                var image = Image.FromFile(mole.ImagePath);
                                picPreview.Image = image;
                            }
                            catch
                            {
                                picPreview.BackColor = Color.LightCoral;
                            }
                        }
                        else if (mole.IsIdleClick)
                        {
                            picPreview.BackColor = Color.LightBlue;
                        }
                        else if (mole.IsJump)
                        {
                            picPreview.BackColor = Color.LightYellow;
                        }
                    }
                }
            };

            // 初始化步骤列表
            if (comboGroup.SelectedIndex >= 0)
            {
                comboGroup_SelectedIndexChanged(null, EventArgs.Empty);
            }

            var hintLabel = new Label
            {
                Text = "提示: 不选择步骤则从分组开始执行；选择步骤则从该步骤开始执行",
                Location = new Point(20, 145),
                Size = new Size(310, 40),
                ForeColor = Color.Gray,
                AutoSize = false,
                Parent = form
            };

            var btnUpdate = new Button
            {
                Text = "更新",
                Location = new Point(100, 310),
                Size = new Size(80, 30),
                Parent = form
            };

            var btnDelete = new Button
            {
                Text = "删除",
                Location = new Point(190, 310),
                Size = new Size(80, 30),
                Parent = form
            };

            var btnCancel = new Button
            {
                Text = "取消",
                Location = new Point(280, 310),
                Size = new Size(80, 30),
                Parent = form
            };
            
            // 更新按钮点击事件
            btnUpdate.Click += (s, e) =>
            {
                if (comboGroup.SelectedIndex >= 0)
                {
                    var targetGroupName = comboGroup.SelectedItem.ToString();
                    var stepIndex = comboStep.SelectedIndex - 1; // -1 表示从头开始
                    
                    // 更新跳转步骤
                    jumpMole.JumpTargetGroup = targetGroupName;
                    jumpMole.JumpTargetStep = stepIndex;
                    jumpMole.Name = stepIndex < 0 
                        ? $"🔗 跳转到 {targetGroupName}" 
                        : $"🔗 跳转到 {targetGroupName} (步骤 {stepIndex + 1})";
                    
                    SaveMoles();
                    
                    // 刷新列表显示
                    var lstMoles = GetCurrentMoleListBox();
                    if (lstMoles != null)
                    {
                        lstMoles.Items[moleIndex] = jumpMole.Name;
                    }
                    
                    var stepInfo = stepIndex < 0 ? "从头开始" : $"从步骤 {stepIndex + 1} 开始";
                    AppendLog($"✅ 已更新跳转步骤: 跳转到 {targetGroupName} ({stepInfo})", LogType.Success);
                    form.Close();
                }
            };
            
            // 取消按钮点击事件
            btnCancel.Click += (s, e) =>
            {
                form.Close();
            };

            // 处理分组选择变化的事件
            void comboGroup_SelectedIndexChanged(object? s, EventArgs e)
            {
                comboStep.Items.Clear();
                comboStep.Items.Add("(从头开始)");
                
                if (comboGroup.SelectedIndex >= 0 && comboGroup.SelectedIndex < otherGroups.Count)
                {
                    var selectedGroup = otherGroups[comboGroup.SelectedIndex];
                    for (int i = 0; i < selectedGroup.Moles.Count; i++)
                    {
                        var mole = selectedGroup.Moles[i];
                        var displayName = mole.IsIdleClick && mole.IdleClickPosition.HasValue
                            ? $"{i + 1}. 💤 {mole.Name}"
                            : mole.IsJump
                            ? $"{i + 1}. 🔗 {mole.Name}"
                            : $"{i + 1}. {mole.Name}";
                        comboStep.Items.Add(displayName);
                    }
                }
                
                // 恢复之前的步骤选择
                if (comboGroup.SelectedIndex >= 0 && comboGroup.SelectedIndex == currentGroupIndex)
                {
                    int stepIndex = jumpMole.JumpTargetStep + 1; // +1 因为第一项是"从头开始"
                    if (stepIndex >= 0 && stepIndex < comboStep.Items.Count)
                        comboStep.SelectedIndex = stepIndex;
                    else
                        comboStep.SelectedIndex = 0;
                }
                else
                {
                    comboStep.SelectedIndex = 0;
                }
            }

            // 删除按钮点击事件
            btnDelete.Click += (s, e) =>
            {
                var result = MessageBox.Show(
                    $"确定要删除跳转步骤 \"{jumpMole.Name}\" 吗？",
                    "确认删除",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Question);

                if (result == DialogResult.Yes)
                {
                    currentGroup.Moles.RemoveAt(moleIndex);
                    SaveMoles();
                    
                    // 刷新列表显示
                    var lstMoles = GetCurrentMoleListBox();
                    if (lstMoles != null)
                    {
                        lstMoles.Items.RemoveAt(moleIndex);
                    }
                    
                    AppendLog($"✅ 已删除跳转步骤: {jumpMole.Name}", LogType.Success);
                    form.Close();
                }
            };

            // 对话框关闭时释放预览图资源和清除引用
            form.FormClosed += (s, e) =>
            {
                if (picPreview.Image != null)
                {
                    var img = picPreview.Image;
                    picPreview.Image = null;
                    img.Dispose();
                }
                
                if (_currentEditDialog == form)
                {
                    _currentEditDialog = null;
                }
            };
            
            // 保存当前编辑窗口引用
            _currentEditDialog = form;
            
            form.Show();
        }

        private void ShowMoleDeleteConfirmDialog(MoleItem mole, int stepIndex)
        {
            var dialog = new Form
            {
                Text = $"步骤 {stepIndex + 1} - 地鼠预览",
                Size = new Size(500, 720),
                StartPosition = FormStartPosition.Manual,
                FormBorderStyle = FormBorderStyle.FixedDialog,
                MaximizeBox = false,
                MinimizeBox = false,
                TopMost = false,
                ShowInTaskbar = false,
                Owner = this
            };
            
            // 设置对话框位置：弹窗左边界与主窗口右边界对齐
            dialog.Location = new Point(
                this.Right,
                this.Top + (this.Height - dialog.Height) / 2
            );
            
            // 提示文字
            var lblMessage = new Label
            {
                Text = $"步骤 {stepIndex + 1}: {mole.Name}",
                Location = new Point(20, 20),
                Size = new Size(350, 30),
                Font = new Font(Font.FontFamily, 10, FontStyle.Bold),
                Parent = dialog
            };
            
            // 预览图
            PictureBox? picPreview = null;
            try
            {
                if (File.Exists(mole.ImagePath))
                {
                    var image = Image.FromFile(mole.ImagePath);
                    
                    // 计算缩放比例，最大显示 300x200
                    int maxWidth = 350;
                    int maxHeight = 200;
                    double scale = Math.Min((double)maxWidth / image.Width, (double)maxHeight / image.Height);
                    if (scale > 1) scale = 1; // 不放大
                    
                    int displayWidth = (int)(image.Width * scale);
                    int displayHeight = (int)(image.Height * scale);
                    
                    picPreview = new PictureBox
                    {
                        Image = image,
                        Location = new Point((dialog.Width - displayWidth) / 2, 60),
                        Size = new Size(displayWidth, displayHeight),
                        SizeMode = PictureBoxSizeMode.Zoom,
                        BorderStyle = BorderStyle.FixedSingle,
                        Parent = dialog
                    };
                    
                    // 显示图像尺寸信息
                    var lblInfo = new Label
                    {
                        Text = $"尺寸: {image.Width} x {image.Height} 像素",
                        Location = new Point(20, picPreview.Bottom + 10),
                        Size = new Size(350, 20),
                        ForeColor = Color.Gray,
                        Parent = dialog
                    };
                }
            }
            catch
            {
                var lblError = new Label
                {
                    Text = "⚠️ 无法加载预览图",
                    Location = new Point(20, 60),
                    Size = new Size(350, 200),
                    TextAlign = System.Drawing.ContentAlignment.MiddleCenter,
                    ForeColor = Color.Red,
                    Parent = dialog
                };
            }
            
            // 匹配阈值标签
            var lblThreshold = new Label
            {
                Text = "匹配阈值 (0.0-1.0):",
                Location = new Point(20, dialog.Height - 400),
                Size = new Size(150, 20),
                Parent = dialog
            };
            
            // 匹配阈值输入框
            var txtThreshold = new TextBox
            {
                Text = mole.SimilarityThreshold.ToString("0.00"),
                Location = new Point(170, dialog.Height - 403),
                Size = new Size(80, 25),
                Parent = dialog
            };
            
            // 阈值说明
            var lblThresholdHint = new Label
            {
                Text = "值越大越严格，默认0.85",
                Location = new Point(260, dialog.Height - 400),
                Size = new Size(120, 20),
                ForeColor = Color.Gray,
                Parent = dialog
            };
            
            // 持续点击直到消失复选框
            var chkClickUntilDisappear = new CheckBox
            {
                Text = "持续点击直到消失",
                Location = new Point(20, dialog.Height - 370),
                Size = new Size(200, 25),
                Checked = mole.ClickUntilDisappear,
                Parent = dialog
            };
            
            // 持续点击说明
            var lblClickHint = new Label
            {
                Text = "识别成功后持续点击，直到图像消失",
                Location = new Point(40, dialog.Height - 345),
                Size = new Size(300, 20),
                ForeColor = Color.Gray,
                Font = new Font(Font.FontFamily, 8),
                Parent = dialog
            };
            
            // 持续等待直到出现复选框
            var chkWaitUntilAppear = new CheckBox
            {
                Text = "持续等待直到出现",
                Location = new Point(20, dialog.Height - 320),
                Size = new Size(200, 25),
                Checked = mole.WaitUntilAppear,
                Parent = dialog
            };
            
            // 持续等待说明
            var lblWaitHint = new Label
            {
                Text = "如果未识别到，重复扫描直到图像出现",
                Location = new Point(40, dialog.Height - 295),
                Size = new Size(300, 20),
                ForeColor = Color.Gray,
                Font = new Font(Font.FontFamily, 8),
                Parent = dialog
            };
            
            // 识别失败跳转到上一步复选框
            var chkJumpToPreviousOnFail = new CheckBox
            {
                Text = "识别失败，跳转到上一个步骤",
                Location = new Point(20, dialog.Height - 270),
                Size = new Size(250, 25),
                Checked = mole.JumpToPreviousOnFail,
                Parent = dialog
            };
            
            // 跳转说明
            var lblJumpHint = new Label
            {
                Text = "未识别到图像时，返回上一步重新执行",
                Location = new Point(40, dialog.Height - 245),
                Size = new Size(300, 20),
                ForeColor = Color.Gray,
                Font = new Font(Font.FontFamily, 8),
                Parent = dialog
            };
            
            // 点击后等待复选框
            var chkWaitAfterClick = new CheckBox
            {
                Text = "成功点击后等待",
                Location = new Point(20, dialog.Height - 220),
                Size = new Size(150, 25),
                Checked = mole.WaitAfterClick,
                Parent = dialog
            };
            
            // 等待时间标签
            var lblWaitTime = new Label
            {
                Text = "等待时间 (ms):",
                Location = new Point(180, dialog.Height - 217),
                Size = new Size(100, 20),
                Parent = dialog
            };
            
            // 等待时间输入框
            var txtWaitTime = new TextBox
            {
                Text = mole.WaitAfterClickMs.ToString(),
                Location = new Point(280, dialog.Height - 220),
                Size = new Size(80, 25),
                Parent = dialog
            };
            
            // 等待说明
            var lblWaitAfterHint = new Label
            {
                Text = "点击成功后等待指定时间再进入下一步",
                Location = new Point(40, dialog.Height - 195),
                Size = new Size(300, 20),
                ForeColor = Color.Gray,
                Font = new Font(Font.FontFamily, 8),
                Parent = dialog
            };
            
            // 按钮
            var btnDelete = new Button
            {
                Text = "删除",
                Location = new Point(dialog.Width / 2 - 220, dialog.Height - 100),
                Size = new Size(80, 30),
                Parent = dialog
            };
            
            var btnCancel = new Button
            {
                Text = "取消",
                Location = new Point(dialog.Width / 2 - 130, dialog.Height - 100),
                Size = new Size(80, 30),
                Parent = dialog
            };
            
            var btnConfirm = new Button
            {
                Text = "确定",
                Location = new Point(dialog.Width / 2 - 40, dialog.Height - 100),
                Size = new Size(80, 30),
                Parent = dialog
            };
            
            var btnUpdateScreenshot = new Button
            {
                Text = "更新截图",
                Location = new Point(dialog.Width / 2 + 50, dialog.Height - 100),
                Size = new Size(80, 30),
                Parent = dialog
            };
            
            // 确定按钮点击事件
            btnConfirm.Click += (s, e) =>
            {
                // 验证并保存阈值
                if (!double.TryParse(txtThreshold.Text, out double threshold))
                {
                    MessageBox.Show("请输入有效的阈值数字", "输入错误", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
                
                if (threshold < 0.0 || threshold > 1.0)
                {
                    MessageBox.Show("阈值必须在 0.0 到 1.0 之间", "输入错误", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
                
                // 验证等待时间
                if (!int.TryParse(txtWaitTime.Text, out int waitTime))
                {
                    MessageBox.Show("请输入有效的等待时间数字", "输入错误", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
                
                if (waitTime < 0)
                {
                    MessageBox.Show("等待时间不能为负数", "输入错误", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
                
                // 保存所有设置
                mole.SimilarityThreshold = threshold;
                mole.ClickUntilDisappear = chkClickUntilDisappear.Checked;
                mole.WaitUntilAppear = chkWaitUntilAppear.Checked;
                mole.JumpToPreviousOnFail = chkJumpToPreviousOnFail.Checked;
                mole.WaitAfterClick = chkWaitAfterClick.Checked;
                mole.WaitAfterClickMs = waitTime;
                SaveMoles();
                AppendLog($"✅ 已更新地鼠 \"{mole.Name}\" 的设置", LogType.Success);
                dialog.Close();
            };
            
            // 更新截图按钮点击事件
            btnUpdateScreenshot.Click += (s, e) =>
            {
                // 先释放预览图资源
                if (picPreview?.Image != null)
                {
                    var img = picPreview.Image;
                    picPreview.Image = null;
                    img.Dispose();
                }
                
                // 关闭当前对话框
                dialog.Close();
                
                // 最小化窗口
                WindowState = FormWindowState.Minimized;
                Thread.Sleep(500);
                
                // 截图
                var screenshot = CaptureScreen();
                
                // 恢复窗口
                WindowState = FormWindowState.Normal;
                
                // 显示截图选择对话框
                var screenshotDialog = new Form
                {
                    Text = "选择新的地鼠区域",
                    Size = new Size(Screen.PrimaryScreen.Bounds.Width, Screen.PrimaryScreen.Bounds.Height),
                    StartPosition = FormStartPosition.Manual,
                    Location = new Point(0, 0),
                    FormBorderStyle = FormBorderStyle.None,
                    WindowState = FormWindowState.Maximized,
                    BackgroundImage = screenshot,
                    BackgroundImageLayout = ImageLayout.Stretch
                };
                
                Point? startPoint = null;
                Rectangle? selection = null;
                
                screenshotDialog.MouseDown += (sd, me) =>
                {
                    if (me.Button == MouseButtons.Left)
                    {
                        startPoint = me.Location;
                    }
                };
                
                screenshotDialog.MouseMove += (sd, me) =>
                {
                    if (startPoint.HasValue)
                    {
                        var rect = GetRectangle(startPoint.Value, me.Location);
                        selection = rect;
                        screenshotDialog.Invalidate();
                    }
                };
                
                screenshotDialog.MouseUp += (sd, me) =>
                {
                    if (me.Button == MouseButtons.Left && selection.HasValue)
                    {
                        screenshotDialog.DialogResult = DialogResult.OK;
                        screenshotDialog.Close();
                    }
                };
                
                screenshotDialog.Paint += (sd, pe) =>
                {
                    if (selection.HasValue)
                    {
                        using (var pen = new Pen(Color.Red, 2))
                        {
                            pe.Graphics.DrawRectangle(pen, selection.Value);
                        }
                    }
                };
                
                screenshotDialog.KeyDown += (sd, ke) =>
                {
                    if (ke.KeyCode == Keys.Escape)
                    {
                        screenshotDialog.DialogResult = DialogResult.Cancel;
                        screenshotDialog.Close();
                    }
                };
                
                if (screenshotDialog.ShowDialog() == DialogResult.OK && selection.HasValue)
                {
                    // 裁剪新图像
                    var croppedImage = CropImage(screenshot, selection.Value);
                    
                    // 删除旧截图文件
                    if (File.Exists(mole.ImagePath))
                    {
                        try
                        {
                            File.Delete(mole.ImagePath);
                        }
                        catch (Exception ex)
                        {
                            AppendLog($"⚠️ 删除旧截图失败: {ex.Message}", LogType.Warning);
                        }
                    }
                    
                    // 保存新截图（使用相同的文件名）
                    croppedImage.Save(mole.ImagePath, System.Drawing.Imaging.ImageFormat.Png);
                    croppedImage.Dispose();
                    
                    SaveMoles();
                    RefreshCurrentMoleList();
                    AppendLog($"✅ 已更新地鼠 \"{mole.Name}\" 的截图", LogType.Success);
                }
                
                screenshot.Dispose();
            };
            
            // 删除按钮点击事件
            btnDelete.Click += (s, e) =>
            {
                // 先释放预览图资源
                if (picPreview?.Image != null)
                {
                    var img = picPreview.Image;
                    picPreview.Image = null;
                    img.Dispose();
                }
                
                // 清空全局预览窗口（如果存在）
                HidePreview();
                
                // 清空图像匹配缓存
                _moleHunter?.ClearTemplateCache();
                
                // 关闭对话框
                dialog.Close();
                
                // 使用异步方式删除，避免阻塞UI
                Task.Run(() =>
                {
                    try
                    {
                        // 等待资源释放
                        System.Threading.Thread.Sleep(300);
                        
                        // 强制垃圾回收
                        GC.Collect();
                        GC.WaitForPendingFinalizers();
                        GC.Collect();
                        
                        // 再等待一下
                        System.Threading.Thread.Sleep(200);
                        
                        // 尝试删除文件（带重试机制）
                        if (!string.IsNullOrEmpty(mole.ImagePath) && File.Exists(mole.ImagePath))
                        {
                            bool deleted = TryDeleteFileWithRetry(mole.ImagePath, maxRetries: 5, delayMs: 500);
                            
                            if (!deleted)
                            {
                                // 删除失败，标记为待删除
                                Invoke(new Action(() =>
                                {
                                    AppendLog($"⚠️ 文件被占用，已标记为待删除: {mole.Name}", LogType.Warning);
                                    AppendLog($"💡 提示: 文件将在下次启动时自动删除", LogType.Info);
                                    
                                    // 标记文件为待删除（下次启动时删除）
                                    MarkFileForDeletion(mole.ImagePath);
                                }));
                            }
                        }
                        
                        // 在UI线程更新界面
                        Invoke(new Action(() =>
                        {
                            // 从当前分组中移除该步骤
                            var group = GetCurrentMoleGroup();
                            var moleToRemove = group.Moles.FirstOrDefault(m => m.ImagePath == mole.ImagePath);
                            if (moleToRemove != null)
                            {
                                group.Moles.Remove(moleToRemove);
                            }
                            
                            // 保存配置
                            SaveMoles();
                            
                            // 刷新列表显示
                            RefreshCurrentMoleList();
                            
                            AppendLog($"✅ 已删除地鼠: {mole.Name}", LogType.Success);
                        }));
                    }
                    catch (Exception ex)
                    {
                        Invoke(new Action(() =>
                        {
                            AppendLog($"❌ 删除失败: {ex.Message}", LogType.Error);
                            MessageBox.Show($"删除失败: {ex.Message}\n\n文件路径: {mole.ImagePath}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }));
                    }
                });
            };
            
            // 取消按钮点击事件
            btnCancel.Click += (s, e) =>
            {
                dialog.Close();
            };
            
            // 注释掉自动关闭功能，改为通过右键切换
            // dialog.Deactivate += (s, e) =>
            // {
            //     if (dialog != null && !dialog.IsDisposed && dialog.Visible)
            //     {
            //         dialog.Close();
            //     }
            // };
            
            // 支持ESC键关闭对话框
            dialog.KeyPreview = true;
            dialog.KeyDown += (s, e) =>
            {
                if (e.KeyCode == Keys.Escape)
                {
                    dialog.Close();
                }
            };
            
            // 对话框关闭时释放图像资源和清除引用
            dialog.FormClosed += (s, e) =>
            {
                if (picPreview?.Image != null)
                {
                    var img = picPreview.Image;
                    picPreview.Image = null;
                    img.Dispose();
                }
                
                if (_currentEditDialog == dialog)
                {
                    _currentEditDialog = null;
                }
            };
            
            // 保存当前编辑窗口引用
            _currentEditDialog = dialog;
            
            // 使用非模态对话框
            dialog.Show();
            
            // 设置焦点到删除按钮
            btnDelete.Focus();
        }
        
        // ==================== 地鼠分组管理 ====================
        
        private void BtnAddMoleGroup_Click(object? sender, EventArgs e)
        {
            var groupName = $"分组 {_moleGroups.Count + 1}";
            var newGroup = new MoleGroup { Name = groupName };
            _moleGroups.Add(newGroup);
            
            CreateMoleGroupTab(newGroup, _moleGroups.Count - 1);
            tabMoleGroups.SelectedIndex = tabMoleGroups.TabPages.Count - 1;
            
            SaveMoles();
            AppendLog($"✅ 已添加新分组: {groupName}", LogType.Success);
        }
        
        private void BtnRemoveMoleGroup_Click(object? sender, EventArgs e)
        {
            if (_moleGroups.Count <= 1)
            {
                MessageBox.Show("至少需要保留一个分组", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            
            var result = MessageBox.Show($"确定要删除分组 \"{_moleGroups[_currentMoleGroupIndex].Name}\" 吗？\n\n该分组下的所有地鼠图片将被删除！", 
                "确认删除", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
            
            if (result == DialogResult.Yes)
            {
                var group = _moleGroups[_currentMoleGroupIndex];
                
                // 删除该组的所有图片文件
                foreach (var mole in group.Moles)
                {
                    if (!mole.IsIdleClick && !string.IsNullOrEmpty(mole.ImagePath) && File.Exists(mole.ImagePath))
                    {
                        try
                        {
                            File.Delete(mole.ImagePath);
                        }
                        catch { }
                    }
                }
                
                _moleGroups.RemoveAt(_currentMoleGroupIndex);
                tabMoleGroups.TabPages.RemoveAt(_currentMoleGroupIndex);
                
                if (_currentMoleGroupIndex >= _moleGroups.Count)
                {
                    _currentMoleGroupIndex = _moleGroups.Count - 1;
                }
                
                if (tabMoleGroups.TabPages.Count > 0)
                {
                    tabMoleGroups.SelectedIndex = _currentMoleGroupIndex;
                }
                
                SaveMoles();
                AppendLog($"✅ 已删除分组: {group.Name}", LogType.Success);
            }
        }
        
        private void TabMoleGroups_SelectedIndexChanged(object? sender, EventArgs e)
        {
            HidePreview(); // 切换标签页时隐藏预览
            
            if (tabMoleGroups.SelectedIndex >= 0)
            {
                _currentMoleGroupIndex = tabMoleGroups.SelectedIndex;
                UpdateIdleClickLabel();
            }
        }
        
        private void TabMoleGroups_MouseDoubleClick(object? sender, MouseEventArgs e)
        {
            // 检查是否双击在标签页标题上
            for (int i = 0; i < tabMoleGroups.TabPages.Count; i++)
            {
                var rect = tabMoleGroups.GetTabRect(i);
                if (rect.Contains(e.Location))
                {
                    // 双击了标签页 i
                    var currentName = _moleGroups[i].Name;
                    var newName = Interaction.InputBox(
                        "请输入新的分组名称:", 
                        "重命名分组", 
                        currentName);
                    
                    if (!string.IsNullOrWhiteSpace(newName) && newName != currentName)
                    {
                        _moleGroups[i].Name = newName;
                        tabMoleGroups.TabPages[i].Text = newName;
                        SaveMoles();
                        AppendLog($"✅ 已重命名分组: {currentName} → {newName}", LogType.Success);
                    }
                    break;
                }
            }
        }

        /// <summary>
        /// 尝试删除文件，带重试机制
        /// </summary>
        private bool TryDeleteFileWithRetry(string filePath, int maxRetries = 5, int delayMs = 500)
        {
            for (int i = 0; i < maxRetries; i++)
            {
                try
                {
                    if (File.Exists(filePath))
                    {
                        File.Delete(filePath);
                        
                        // 验证是否真的删除了
                        if (!File.Exists(filePath))
                        {
                            return true;
                        }
                    }
                    else
                    {
                        // 文件不存在，认为删除成功
                        return true;
                    }
                }
                catch (IOException)
                {
                    // 文件被占用，等待后重试
                    if (i < maxRetries - 1)
                    {
                        System.Threading.Thread.Sleep(delayMs);
                    }
                }
                catch (UnauthorizedAccessException)
                {
                    // 权限问题，等待后重试
                    if (i < maxRetries - 1)
                    {
                        System.Threading.Thread.Sleep(delayMs);
                    }
                }
            }
            
            return false;
        }

        /// <summary>
        /// 标记文件为待删除（下次启动时删除）
        /// </summary>
        private void MarkFileForDeletion(string filePath)
        {
            try
            {
                var pendingDeleteFile = Path.Combine(_molesDirectory, "pending_delete.txt");
                File.AppendAllText(pendingDeleteFile, filePath + Environment.NewLine);
            }
            catch
            {
                // 忽略错误
            }
        }

        /// <summary>
        /// 处理待删除的文件（启动时调用）
        /// </summary>
        private void ProcessPendingDeletions()
        {
            try
            {
                var pendingDeleteFile = Path.Combine(_molesDirectory, "pending_delete.txt");
                
                if (!File.Exists(pendingDeleteFile))
                    return;
                
                var filesToDelete = File.ReadAllLines(pendingDeleteFile)
                    .Where(line => !string.IsNullOrWhiteSpace(line))
                    .ToList();
                
                var deletedFiles = new List<string>();
                
                foreach (var filePath in filesToDelete)
                {
                    if (File.Exists(filePath))
                    {
                        try
                        {
                            File.Delete(filePath);
                            deletedFiles.Add(filePath);
                            AppendLog($"✅ 已删除待删除文件: {Path.GetFileName(filePath)}", LogType.Success);
                        }
                        catch
                        {
                            // 仍然无法删除，保留在列表中
                        }
                    }
                    else
                    {
                        // 文件已不存在，从列表中移除
                        deletedFiles.Add(filePath);
                    }
                }
                
                // 更新待删除列表
                var remainingFiles = filesToDelete.Except(deletedFiles).ToList();
                
                if (remainingFiles.Count > 0)
                {
                    File.WriteAllLines(pendingDeleteFile, remainingFiles);
                }
                else
                {
                    // 所有文件都已删除，删除待删除列表文件
                    File.Delete(pendingDeleteFile);
                }
            }
            catch
            {
                // 忽略错误
            }
        }
    }
}
