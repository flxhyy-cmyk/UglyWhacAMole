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
            try
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
            catch (Exception ex)
            {
                MessageBox.Show($"初始化失败: {ex.Message}\n\n{ex.StackTrace}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                throw;
            }
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
                        AppendLog($"🔄 正在加载上次的配置: {lastConfigName}", LogType.Info);
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
                            AppendLog($"✅ 已自动加载配置: {lastConfigName}", LogType.Success);
                        }
                        else
                        {
                            AppendLog($"⚠️ 配置文件解析失败: {lastConfigName}", LogType.Warning);
                        }
                    }
                    catch (Exception ex)
                    {
                        AppendLog($"❌ 加载配置失败: {lastConfigName}", LogType.Error);
                        AppendLog($"错误详情: {ex.Message}", LogType.Error);
                    }
                }
                else
                {
                    AppendLog($"⚠️ 上次的配置文件不存在: {lastConfigName}", LogType.Warning);
                }
            }
            else
            {
                AppendLog("ℹ️ 没有上次使用的配置记录", LogType.Info);
                // 加载默认配置
                var config = _configManager.LoadConfig();
                if (config != null)
                {
                    AppendLog("🔄 正在加载默认配置", LogType.Info);
                    _config = config;
                    UpdateTextCombo();
                    UpdateCellGroupCombo();
                    TryAutoFindWindow();
                    AppendLog("✅ 已加载默认配置", LogType.Success);
                }
                else
                {
                    AppendLog("ℹ️ 没有默认配置，使用空配置", LogType.Info);
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
            _moleHunter.OnConfigSwitchRequested += MoleHunter_OnConfigSwitchRequested;
            _moleHunter.OnTextContentSwitchRequested += MoleHunter_OnTextContentSwitchRequested;
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
            TryAutoFindWindow(true);
        }

        private void TryAutoFindWindow(bool isStartup)
        {
            if (string.IsNullOrEmpty(_config.WindowClass))
            {
                if (!isStartup)
                {
                    AppendLog("⚠️ 配置中没有窗口类名，无法查找窗口", LogType.Warning);
                }
                return;
            }

            if (!isStartup)
            {
                AppendLog($"🔍 正在查找目标窗口...", LogType.Info);
            }

            IntPtr foundWindow = FindTargetWindow();

            if (foundWindow != IntPtr.Zero)
            {
                _targetWindow = foundWindow;
                WindowHelper.GetWindowRect(_targetWindow, out _windowRect);
                
                if (isStartup)
                {
                    AppendLog($"✅ 成功找到目标窗口 (句柄: 0x{_targetWindow.ToInt64():X})", LogType.Success);
                    OnWindowSelected(_config.WindowTitle, true);
                }
                else
                {
                    AppendLog($"✅ 成功找到目标窗口 (句柄: 0x{_targetWindow.ToInt64():X})", LogType.Success);
                }
            }
            else
            {
                if (isStartup)
                {
                    // 启动时静默处理，只显示温和提示
                    AppendLog($"ℹ️ 目标窗口暂未运行，将在填充时自动查找", LogType.Info);
                    ShowLoadedConfigInfo();
                }
                else
                {
                    // 填充时未找到才明确提示
                    AppendLog($"❌ 未找到目标窗口", LogType.Error);
                    if (!string.IsNullOrEmpty(_config.TargetProgramPath))
                    {
                        AppendLog($"   目标程序路径: {_config.TargetProgramPath}", LogType.Info);
                    }
                }
            }
        }

        private IntPtr FindTargetWindow()
        {
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

            return foundWindow;
        }

        private bool IsWindowValid(IntPtr hwnd)
        {
            if (hwnd == IntPtr.Zero)
                return false;

            // 检查窗口是否仍然存在
            try
            {
                var className = WindowHelper.GetWindowClassName(hwnd);
                return !string.IsNullOrEmpty(className);
            }
            catch
            {
                return false;
            }
        }

        private bool EnsureTargetWindowValid()
        {
            // 如果窗口句柄有效，直接返回
            if (IsWindowValid(_targetWindow))
                return true;

            // 窗口句柄无效，尝试重新查找
            AppendLog($"🔍 正在查找目标窗口...", LogType.Info);
            
            IntPtr foundWindow = FindTargetWindow();
            
            if (foundWindow != IntPtr.Zero)
            {
                _targetWindow = foundWindow;
                WindowHelper.GetWindowRect(_targetWindow, out _windowRect);
                AppendLog($"✅ 成功找到目标窗口 (句柄: 0x{_targetWindow.ToInt64():X})", LogType.Success);
                return true;
            }
            
            // 仍然找不到，给出明确提示
            AppendLog($"❌ 未找到目标窗口", LogType.Error);
            AppendLog($"   窗口类名: {_config.WindowClass}", LogType.Info);
            if (!_config.IsExcelMode)
            {
                AppendLog($"   窗口标题: {_config.WindowTitle}", LogType.Info);
            }
            
            if (!string.IsNullOrEmpty(_config.TargetProgramPath))
            {
                AppendLog($"   请先启动: {Path.GetFileName(_config.TargetProgramPath)}", LogType.Warning);
            }
            else
            {
                AppendLog($"   请先启动目标程序", LogType.Warning);
            }
            
            return false;
        }

        private void ShowLoadedConfigInfo()
        {
            AppendLog($"\n📋 已加载配置信息:", LogType.Info);
            
            if (_config.IsExcelMode)
            {
                AppendLog("📊 模式: Excel专用模式", LogType.Info);
                if (_config.ExcelCells.Count > 0)
                {
                    AppendLog($"   Excel单元格数量: {_config.ExcelCells.Count}", LogType.Info);
                    AppendLog($"   单元格地址: {string.Join(", ", _config.ExcelCells)}", LogType.Info);
                }
            }
            else
            {
                AppendLog("📝 模式: 普通窗口模式", LogType.Info);
                if (_config.InputPositions.Count > 0)
                {
                    AppendLog($"   输入框位置数量: {_config.InputPositions.Count}", LogType.Info);
                    for (int i = 0; i < _config.InputPositions.Count; i++)
                    {
                        var pos = _config.InputPositions[i];
                        AppendLog($"   输入框 {i + 1}: 相对位置 ({pos.X}, {pos.Y})", LogType.Info);
                    }
                }
            }
            
            if (_config.SavedTexts.Count > 0)
            {
                AppendLog($"   已保存文本数量: {_config.SavedTexts.Count}", LogType.Info);
            }
            
            AppendLog($"\n💡 提示: 启动目标程序后，直接按 F2 即可自动填充", LogType.Info);
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
                    // 填充前确保窗口句柄有效
                    if (!EnsureTargetWindowValid())
                    {
                        return; // 窗口未找到，已在方法内提示用户
                    }
                    
                    await _textFiller.FillTextAsync(_targetWindow, _windowRect, _config.InputPositions, item.Texts);
                }
                
                AppendLog($"✅ 已填充: {item.Name}", LogType.Success);
                
                // 纯粹的顺序跳转：填充完当前项后，跳转到下一个项（循环）
                // 不记录、不判断、只服从用户当前选择
                int nextIndex = (currentIndex + 1) % _config.SavedTexts.Count;
                cmbSavedTexts.SelectedIndex = nextIndex;
                AppendLog($"⏭️ 跳转到: {_config.SavedTexts[nextIndex].Name}", LogType.Info);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"填充失败: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
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
                        
                        // 数据迁移：将旧的IdleClickPositions转换为Moles中的空击步骤
                        bool needsMigration = false;
                        foreach (var group in _moleGroups)
                        {
                            // 检查是否有旧的IdleClickPositions数据（通过反射或尝试反序列化）
                            // 由于我们已经移除了IdleClickPositions字段，这里需要特殊处理
                            // 我们可以尝试从JSON中读取IdleClickPositions
                            try
                            {
                                var jsonToken = Newtonsoft.Json.Linq.JToken.Parse(json);
                                var groupsArray = jsonToken as Newtonsoft.Json.Linq.JArray ?? (jsonToken as Newtonsoft.Json.Linq.JObject)?["$values"] as Newtonsoft.Json.Linq.JArray;
                                
                                if (groupsArray != null)
                                {
                                    for (int i = 0; i < groupsArray.Count && i < _moleGroups.Count; i++)
                                    {
                                        var groupObj = groupsArray[i] as Newtonsoft.Json.Linq.JObject;
                                        if (groupObj != null && groupObj["IdleClickPositions"] != null)
                                        {
                                            var idleClickPositions = groupObj["IdleClickPositions"].ToObject<List<Point>>();
                                            if (idleClickPositions != null && idleClickPositions.Count > 0)
                                            {
                                                // 检查是否已经有对应的空击步骤
                                                var existingIdleClicks = _moleGroups[i].Moles.Where(m => m.IsIdleClick).ToList();
                                                
                                                // 只迁移那些不在Moles列表中的空击位置
                                                foreach (var pos in idleClickPositions)
                                                {
                                                    bool exists = existingIdleClicks.Any(m => 
                                                        m.IdleClickPosition.HasValue && 
                                                        m.IdleClickPosition.Value.X == pos.X && 
                                                        m.IdleClickPosition.Value.Y == pos.Y);
                                                    
                                                    if (!exists)
                                                    {
                                                        int idleClickCount = _moleGroups[i].Moles.Count(m => m.IsIdleClick) + 1;
                                                        var idleMole = new MoleItem
                                                        {
                                                            Name = $"空击 {idleClickCount}",
                                                            ImagePath = "",
                                                            IsEnabled = true,
                                                            IsIdleClick = true,
                                                            IdleClickPosition = pos
                                                        };
                                                        _moleGroups[i].Moles.Add(idleMole);
                                                        needsMigration = true;
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                            catch
                            {
                                // 忽略迁移错误
                            }
                        }
                        
                        if (needsMigration)
                        {
                            AppendLog("🔄 检测到旧版本数据，已自动迁移空击位置", LogType.Info);
                            SaveMoles(); // 保存迁移后的数据
                        }
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
            
            // 初始化显示设置界面
            try
            {
                LoadMoleGroupsSelection();
            }
            catch (Exception ex)
            {
                AppendLog($"⚠️ 加载分组选择界面失败: {ex.Message}", LogType.Warning);
            }
            
            // 根据配置决定是否自动显示分组
            if (_config.AutoLoadMoleGroups)
            {
                // 启用了自动显示，显示选中的分组
                if (_config.SelectedMoleGroups.Count > 0)
                {
                    LoadSelectedMoleGroups();
                    AppendLog($"📂 已自动显示 {tabMoleGroups.TabPages.Count} 个选中的分组", LogType.Info);
                }
                else
                {
                    // 没有选中任何分组，默认显示第一个
                    if (_moleGroups.Count > 0)
                    {
                        CreateMoleGroupTab(_moleGroups[0], 0);
                        tabMoleGroups.SelectedIndex = 0;
                        _currentMoleGroupIndex = 0;
                        AppendLog($"📂 已自动显示默认分组", LogType.Info);
                    }
                }
            }
            else
            {
                // 未启用自动显示，不显示任何分组到标签页
                // 用户需要手动在"显示设置"界面点击"显示选中的分组"按钮
                AppendLog($"ℹ️ 已加载 {_moleGroups.Count} 个地鼠分组配置", LogType.Info);
                AppendLog($"💡 请在【显示设置】标签页选择要显示的分组", LogType.Info);
            }
            
            UpdateIdleClickLabel();
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
                
                if (mole.IsConfigStep)
                {
                    displayText = $"{i + 1}. {mole.Name}";
                }
                else if (mole.IsIdleClick && mole.IdleClickPosition.HasValue)
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
            lstMoles.ItemCheck += LstMoles_ItemCheck;
            
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
        
        private void UpdateIdleClickLabel()
        {
            var group = GetCurrentMoleGroup();
            int idleClickCount = group.Moles.Count(m => m.IsIdleClick);
            
            if (idleClickCount > 0)
            {
                lblIdleClickPos.Text = $"空击: {idleClickCount} 个位置";
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
                
                _moleHunter.Start(group.Moles, _moleGroups);
                AppendLog($"🎯 打地鼠已启动 - 分组: {group.Name}", LogType.Success);
                
                int idleClickCount = group.Moles.Count(m => m.IsIdleClick);
                if (idleClickCount > 0)
                {
                    AppendLog($"💤 空击位置数量: {idleClickCount}", LogType.Info);
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
        
        private void MoleHunter_OnConfigSwitchRequested(object? sender, string configName)
        {
            // 在UI线程上执行配置切换
            if (InvokeRequired)
            {
                Invoke(new Action(() => MoleHunter_OnConfigSwitchRequested(sender, configName)));
                return;
            }
            
            try
            {
                LoadNamedConfig(configName);
            }
            catch (Exception ex)
            {
                AppendLog($"❌ 配置切换失败: {ex.Message}", LogType.Error);
            }
        }
        
        private void MoleHunter_OnTextContentSwitchRequested(object? sender, string textName)
        {
            // 在UI线程上执行填充内容切换
            if (InvokeRequired)
            {
                Invoke(new Action(() => MoleHunter_OnTextContentSwitchRequested(sender, textName)));
                return;
            }
            
            try
            {
                // 查找目标文本项
                var targetIndex = _config.SavedTexts.FindIndex(t => t.Name == textName);
                if (targetIndex >= 0)
                {
                    cmbSavedTexts.SelectedIndex = targetIndex;
                }
                else
                {
                    AppendLog($"⚠️ 未找到填充内容: {textName}", LogType.Warning);
                }
            }
            catch (Exception ex)
            {
                AppendLog($"❌ 填充内容切换失败: {ex.Message}", LogType.Error);
            }
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
                        
                        // 计算空击步骤的编号
                        int idleClickCount = group.Moles.Count(m => m.IsIdleClick) + 1;
                        
                        // 直接创建空击步骤并添加到列表末尾
                        var idleMole = new MoleItem
                        {
                            Name = $"空击 {idleClickCount}",
                            ImagePath = "",
                            IsEnabled = true,
                            IsIdleClick = true,
                            IdleClickPosition = newPoint
                        };
                        
                        group.Moles.Add(idleMole);
                        
                        Invoke(new Action(() =>
                        {
                            UpdateIdleClickLabel();
                            AppendLog($"✅ 空击位置 {idleClickCount}: ({pos.X}, {pos.Y})", LogType.Success);
                            RefreshCurrentMoleList();
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
        
        private void BtnAddConfigStep_Click(object? sender, EventArgs e)
        {
            var currentGroup = GetCurrentMoleGroup();
            if (currentGroup == null)
                return;
            
            ShowConfigStepDialog(null, -1);
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

            // 创建选择框（加高窗口以容纳新功能）
            var form = new Form
            {
                Text = "选择跳转目标",
                Size = new Size(350, 620),
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

            // 分隔线
            var separator = new Label
            {
                Text = "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━",
                Location = new Point(20, 190),
                Size = new Size(310, 20),
                ForeColor = Color.Gray,
                Parent = form
            };

            // 键盘按键输入复选框
            var chkSendKeyPress = new CheckBox
            {
                Text = "发送键盘按键输入（忽略跳转逻辑）",
                Location = new Point(20, 215),
                Size = new Size(310, 25),
                Parent = form
            };

            var labelKeyPress = new Label
            {
                Text = "按键定义（点击文本框后按下按键）:",
                Location = new Point(20, 245),
                Size = new Size(310, 20),
                Enabled = false,
                Parent = form
            };

            var txtKeyPress = new TextBox
            {
                Location = new Point(20, 270),
                Size = new Size(310, 25),
                ReadOnly = true,
                Enabled = false,
                PlaceholderText = "点击后按下按键...",
                Parent = form
            };

            var labelWaitTime = new Label
            {
                Text = "按键输入后等待时间（毫秒）:",
                Location = new Point(20, 305),
                Size = new Size(310, 20),
                Enabled = false,
                Parent = form
            };

            var txtWaitTime = new TextBox
            {
                Text = "100",
                Location = new Point(20, 330),
                Size = new Size(310, 25),
                Enabled = false,
                Parent = form
            };

            // 鼠标滚动复选框
            var chkMouseScroll = new CheckBox
            {
                Text = "鼠标滚动操作",
                Location = new Point(20, 365),
                Size = new Size(310, 25),
                Enabled = false,
                Parent = form
            };

            var labelScrollDirection = new Label
            {
                Text = "滚动方向:",
                Location = new Point(40, 395),
                Size = new Size(70, 20),
                Enabled = false,
                Parent = form
            };

            var comboScrollDirection = new ComboBox
            {
                Location = new Point(110, 392),
                Size = new Size(90, 25),
                DropDownStyle = ComboBoxStyle.DropDownList,
                Enabled = false,
                Parent = form
            };
            comboScrollDirection.Items.Add("向上滚动");
            comboScrollDirection.Items.Add("向下滚动");
            comboScrollDirection.SelectedIndex = 0;

            var labelScrollCount = new Label
            {
                Text = "滚动次数:",
                Location = new Point(40, 425),
                Size = new Size(70, 20),
                Enabled = false,
                Parent = form
            };

            var txtScrollCount = new TextBox
            {
                Text = "1",
                Location = new Point(40, 450),
                Size = new Size(260, 25),
                Enabled = false,
                Parent = form
            };

            var labelScrollWait = new Label
            {
                Text = "滚动后延时(ms):",
                Location = new Point(40, 480),
                Size = new Size(110, 20),
                Enabled = false,
                Parent = form
            };

            var txtScrollWait = new TextBox
            {
                Text = "100",
                Location = new Point(40, 505),
                Size = new Size(260, 25),
                Enabled = false,
                Parent = form
            };

            // 复选框状态改变事件
            chkSendKeyPress.CheckedChanged += (s, e) =>
            {
                bool enabled = chkSendKeyPress.Checked;
                labelKeyPress.Enabled = enabled;
                txtKeyPress.Enabled = enabled;
                labelWaitTime.Enabled = enabled;
                txtWaitTime.Enabled = enabled;
                chkMouseScroll.Enabled = enabled;
                
                // 如果禁用按键输入，同时禁用鼠标滚动
                if (!enabled)
                {
                    chkMouseScroll.Checked = false;
                }
                
                // 禁用/启用跳转相关控件
                label1.Enabled = !enabled;
                comboGroup.Enabled = !enabled;
                label2.Enabled = !enabled;
                comboStep.Enabled = !enabled;
            };

            // 鼠标滚动复选框状态改变事件
            chkMouseScroll.CheckedChanged += (s, e) =>
            {
                bool enabled = chkMouseScroll.Checked;
                labelScrollDirection.Enabled = enabled;
                comboScrollDirection.Enabled = enabled;
                labelScrollCount.Enabled = enabled;
                txtScrollCount.Enabled = enabled;
                labelScrollWait.Enabled = enabled;
                txtScrollWait.Enabled = enabled;
            };

            // 按键录制逻辑
            string recordedKey = "";
            bool hotkeysUnregistered = false;
            
            txtKeyPress.Enter += (s, e) =>
            {
                txtKeyPress.Text = "按下按键...";
                recordedKey = "";
                
                // 暂时注销全局热键，允许用户录制 F2、F3、F4、F6
                UnregisterGlobalHotKeys();
                hotkeysUnregistered = true;
            };

            txtKeyPress.Leave += (s, e) =>
            {
                // 恢复全局热键
                if (hotkeysUnregistered)
                {
                    RegisterGlobalHotKeys();
                    hotkeysUnregistered = false;
                }
            };

            txtKeyPress.KeyDown += (s, e) =>
            {
                e.SuppressKeyPress = true; // 阻止默认行为
                
                // 构建按键字符串
                var keyParts = new List<string>();
                
                if (e.Control) keyParts.Add("Ctrl");
                if (e.Shift) keyParts.Add("Shift");
                if (e.Alt) keyParts.Add("Alt");
                
                // 获取主键
                var mainKey = e.KeyCode.ToString();
                
                // 排除修饰键本身
                if (mainKey != "ControlKey" && mainKey != "ShiftKey" && mainKey != "Menu")
                {
                    keyParts.Add(mainKey);
                }
                
                if (keyParts.Count > 0)
                {
                    recordedKey = string.Join("+", keyParts);
                    txtKeyPress.Text = recordedKey;
                }
            };
            
            // 对话框关闭时确保恢复热键
            form.FormClosing += (s, e) =>
            {
                if (hotkeysUnregistered)
                {
                    RegisterGlobalHotKeys();
                    hotkeysUnregistered = false;
                }
            };

            var btnOk = new Button
            {
                Text = "确定",
                Location = new Point(150, 545),
                Size = new Size(80, 30),
                DialogResult = DialogResult.OK,
                Parent = form
            };

            var btnCancel = new Button
            {
                Text = "取消",
                Location = new Point(240, 545),
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

            if (form.ShowDialog() == DialogResult.OK)
            {
                MoleItem jumpMole;
                
                if (chkSendKeyPress.Checked)
                {
                    // 键盘按键输入模式
                    if (string.IsNullOrEmpty(recordedKey))
                    {
                        MessageBox.Show("请先录制按键", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return;
                    }
                    
                    if (!int.TryParse(txtWaitTime.Text, out int waitMs) || waitMs < 0)
                    {
                        MessageBox.Show("等待时间必须是非负整数", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return;
                    }
                    
                    // 验证鼠标滚动参数
                    int scrollCount = 1;
                    int scrollWaitMs = 100;
                    if (chkMouseScroll.Checked)
                    {
                        if (!int.TryParse(txtScrollCount.Text, out scrollCount) || scrollCount < 1)
                        {
                            MessageBox.Show("滚动次数必须是正整数", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            return;
                        }
                        
                        if (!int.TryParse(txtScrollWait.Text, out scrollWaitMs) || scrollWaitMs < 0)
                        {
                            MessageBox.Show("滚动后延时必须是非负整数", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            return;
                        }
                    }
                    
                    jumpMole = new MoleItem
                    {
                        Name = $"⌨️ 按键: {recordedKey}",
                        IsJump = true,
                        SendKeyPress = true,
                        KeyPressDefinition = recordedKey,
                        KeyPressWaitMs = waitMs,
                        EnableMouseScroll = chkMouseScroll.Checked,
                        ScrollUp = comboScrollDirection.SelectedIndex == 0,
                        ScrollCount = scrollCount,
                        ScrollWaitMs = scrollWaitMs,
                        IsEnabled = true
                    };
                    
                    currentGroup.Moles.Add(jumpMole);
                    SaveMoles();
                    
                    var lstMoles = GetCurrentMoleListBox();
                    if (lstMoles != null)
                    {
                        int index = currentGroup.Moles.Count - 1;
                        string displayText = $"{index + 1}. 🔗 {jumpMole.Name}";
                        lstMoles.Items.Add(displayText, true);
                    }
                    
                    var logMsg = $"✅ 已添加按键步骤: {recordedKey} (等待 {waitMs}ms)";
                    if (chkMouseScroll.Checked)
                    {
                        var direction = comboScrollDirection.SelectedIndex == 0 ? "向上" : "向下";
                        logMsg += $" + 鼠标{direction}滚动{scrollCount}次 (延时 {scrollWaitMs}ms)";
                    }
                    AppendLog(logMsg, LogType.Success);
                }
                else
                {
                    // 跳转模式
                    if (comboGroup.SelectedIndex < 0)
                    {
                        MessageBox.Show("请选择跳转目标分组", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return;
                    }
                    
                    var targetGroupName = comboGroup.SelectedItem.ToString();
                    var stepIndex = comboStep.SelectedIndex - 1; // -1 表示从头开始
                    
                    jumpMole = new MoleItem
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
                    
                    var lstMoles = GetCurrentMoleListBox();
                    if (lstMoles != null)
                    {
                        int index = currentGroup.Moles.Count - 1;
                        string displayText = $"{index + 1}. 🔗 {jumpMole.Name}";
                        lstMoles.Items.Add(displayText, true);
                    }

                    var stepInfo = stepIndex < 0 ? "从头开始" : $"从步骤 {stepIndex + 1} 开始";
                    AppendLog($"✅ 已添加跳转步骤: 跳转到 {targetGroupName} ({stepInfo})", LogType.Success);
                }
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
        
        private void LstMoles_ItemCheck(object? sender, ItemCheckEventArgs e)
        {
            if (sender is CheckedListBox lstMoles)
            {
                var group = GetCurrentMoleGroup();
                if (group == null || e.Index < 0 || e.Index >= group.Moles.Count)
                    return;
                
                // 使用 BeginInvoke 延迟执行，因为 ItemCheck 事件在状态实际改变之前触发
                this.BeginInvoke(new Action(() =>
                {
                    // 同步复选框状态到配置
                    group.Moles[e.Index].IsEnabled = lstMoles.GetItemChecked(e.Index);
                    
                    // 实时保存配置
                    SaveMoles();
                    
                    var statusText = group.Moles[e.Index].IsEnabled ? "已启用" : "已禁用";
                    AppendLog($"✅ 步骤 {e.Index + 1} {statusText}: {group.Moles[e.Index].Name}", LogType.Info);
                }));
            }
        }
        
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
                else if (mole.IsConfigStep)
                {
                    displayText = $"{i + 1}. {mole.Name}";
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
                    
                    // 如果是配置步骤，显示编辑对话框
                    if (mole.IsConfigStep)
                    {
                        ShowConfigStepDialog(mole, index);
                        return;
                    }
                    
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
                Location = new Point(190, 170),
                Size = new Size(80, 30),
                Parent = form
            };
            
            var btnDelete = new Button
            {
                Text = "删除",
                Location = new Point(100, 170),
                Size = new Size(80, 30),
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
                    // 直接从Moles列表中移除
                    currentGroup.Moles.Remove(idleMole);
                    
                    AppendLog($"✅ 已删除空击位置: {idleMole.Name}", LogType.Success);
                    RefreshCurrentMoleList();
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
            
            // 自动聚焦删除按钮
            btnDelete.Focus();
        }
        
        private void ShowJumpStepEditDialog(MoleItem jumpMole, int moleIndex)
        {
            var currentGroup = GetCurrentMoleGroup();
            var otherGroups = _moleGroups
                .Where(g => g.Name != currentGroup.Name)
                .ToList();

            if (otherGroups.Count == 0 && !jumpMole.SendKeyPress)
            {
                MessageBox.Show("没有其他分组可以跳转到", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            // 创建编辑对话框（加高以容纳按键输入和鼠标滚动UI）
            var form = new Form
            {
                Text = "编辑跳转步骤",
                Size = new Size(500, 680),
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

            // 分隔线
            var separator = new Label
            {
                Text = "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━",
                Location = new Point(20, 190),
                Size = new Size(310, 20),
                ForeColor = Color.Gray,
                Parent = form
            };

            // 键盘按键输入复选框
            var chkSendKeyPress = new CheckBox
            {
                Text = "发送键盘按键输入（忽略跳转逻辑）",
                Location = new Point(20, 215),
                Size = new Size(310, 25),
                Checked = jumpMole.SendKeyPress,
                Parent = form
            };

            var labelKeyPress = new Label
            {
                Text = "按键定义（点击文本框后按下按键）:",
                Location = new Point(20, 245),
                Size = new Size(310, 20),
                Enabled = jumpMole.SendKeyPress,
                Parent = form
            };

            var txtKeyPress = new TextBox
            {
                Location = new Point(20, 270),
                Size = new Size(310, 25),
                ReadOnly = true,
                Enabled = jumpMole.SendKeyPress,
                Text = jumpMole.KeyPressDefinition,
                PlaceholderText = "点击后按下按键...",
                Parent = form
            };

            var labelWaitTime = new Label
            {
                Text = "按键输入后等待时间（毫秒）:",
                Location = new Point(20, 305),
                Size = new Size(310, 20),
                Enabled = jumpMole.SendKeyPress,
                Parent = form
            };

            var txtWaitTime = new TextBox
            {
                Text = jumpMole.KeyPressWaitMs.ToString(),
                Location = new Point(20, 330),
                Size = new Size(310, 25),
                Enabled = jumpMole.SendKeyPress,
                Parent = form
            };

            // 鼠标滚动复选框
            var chkMouseScroll = new CheckBox
            {
                Text = "鼠标滚动操作",
                Location = new Point(20, 365),
                Size = new Size(310, 25),
                Checked = jumpMole.EnableMouseScroll,
                Enabled = jumpMole.SendKeyPress,
                Parent = form
            };

            var labelScrollDirection = new Label
            {
                Text = "滚动方向:",
                Location = new Point(40, 395),
                Size = new Size(70, 20),
                Enabled = jumpMole.EnableMouseScroll,
                Parent = form
            };

            var comboScrollDirection = new ComboBox
            {
                Location = new Point(110, 392),
                Size = new Size(90, 25),
                DropDownStyle = ComboBoxStyle.DropDownList,
                Enabled = jumpMole.EnableMouseScroll,
                Parent = form
            };
            comboScrollDirection.Items.Add("向上滚动");
            comboScrollDirection.Items.Add("向下滚动");
            comboScrollDirection.SelectedIndex = jumpMole.ScrollUp ? 0 : 1;

            var labelScrollCount = new Label
            {
                Text = "滚动次数:",
                Location = new Point(40, 425),
                Size = new Size(70, 20),
                Enabled = jumpMole.EnableMouseScroll,
                Parent = form
            };

            var txtScrollCount = new TextBox
            {
                Text = jumpMole.ScrollCount.ToString(),
                Location = new Point(40, 450),
                Size = new Size(260, 25),
                Enabled = jumpMole.EnableMouseScroll,
                Parent = form
            };

            var labelScrollWait = new Label
            {
                Text = "滚动后延时(ms):",
                Location = new Point(40, 480),
                Size = new Size(110, 20),
                Enabled = jumpMole.EnableMouseScroll,
                Parent = form
            };

            var txtScrollWait = new TextBox
            {
                Text = jumpMole.ScrollWaitMs.ToString(),
                Location = new Point(40, 505),
                Size = new Size(260, 25),
                Enabled = jumpMole.EnableMouseScroll,
                Parent = form
            };

            // 复选框状态改变事件
            chkSendKeyPress.CheckedChanged += (s, e) =>
            {
                bool enabled = chkSendKeyPress.Checked;
                labelKeyPress.Enabled = enabled;
                txtKeyPress.Enabled = enabled;
                labelWaitTime.Enabled = enabled;
                txtWaitTime.Enabled = enabled;
                chkMouseScroll.Enabled = enabled;
                
                // 如果禁用按键输入，同时禁用鼠标滚动
                if (!enabled)
                {
                    chkMouseScroll.Checked = false;
                }
                
                // 禁用/启用跳转相关控件
                label1.Enabled = !enabled;
                comboGroup.Enabled = !enabled;
                label2.Enabled = !enabled;
                comboStep.Enabled = !enabled;
            };

            // 鼠标滚动复选框状态改变事件
            chkMouseScroll.CheckedChanged += (s, e) =>
            {
                bool enabled = chkMouseScroll.Checked;
                labelScrollDirection.Enabled = enabled;
                comboScrollDirection.Enabled = enabled;
                labelScrollCount.Enabled = enabled;
                txtScrollCount.Enabled = enabled;
                labelScrollWait.Enabled = enabled;
                txtScrollWait.Enabled = enabled;
            };

            // 按键录制逻辑
            string recordedKey = jumpMole.KeyPressDefinition;
            bool hotkeysUnregistered = false;
            
            txtKeyPress.Enter += (s, e) =>
            {
                txtKeyPress.Text = "按下按键...";
                recordedKey = "";
                
                // 暂时注销全局热键
                UnregisterGlobalHotKeys();
                hotkeysUnregistered = true;
            };

            txtKeyPress.Leave += (s, e) =>
            {
                // 恢复全局热键
                if (hotkeysUnregistered)
                {
                    RegisterGlobalHotKeys();
                    hotkeysUnregistered = false;
                }
                
                // 如果没有录制到按键，恢复原值
                if (string.IsNullOrEmpty(recordedKey))
                {
                    txtKeyPress.Text = jumpMole.KeyPressDefinition;
                    recordedKey = jumpMole.KeyPressDefinition;
                }
            };

            txtKeyPress.KeyDown += (s, e) =>
            {
                e.SuppressKeyPress = true;
                
                var keyParts = new List<string>();
                
                if (e.Control) keyParts.Add("Ctrl");
                if (e.Shift) keyParts.Add("Shift");
                if (e.Alt) keyParts.Add("Alt");
                
                var mainKey = e.KeyCode.ToString();
                
                if (mainKey != "ControlKey" && mainKey != "ShiftKey" && mainKey != "Menu")
                {
                    keyParts.Add(mainKey);
                }
                
                if (keyParts.Count > 0)
                {
                    recordedKey = string.Join("+", keyParts);
                    txtKeyPress.Text = recordedKey;
                }
            };
            
            // 对话框关闭时确保恢复热键
            form.FormClosing += (s, e) =>
            {
                if (hotkeysUnregistered)
                {
                    RegisterGlobalHotKeys();
                    hotkeysUnregistered = false;
                }
            };

            var btnUpdate = new Button
            {
                Text = "更新",
                Location = new Point(100, 610),
                Size = new Size(80, 30),
                Parent = form
            };

            var btnDelete = new Button
            {
                Text = "删除",
                Location = new Point(190, 610),
                Size = new Size(80, 30),
                Parent = form
            };

            var btnCancel = new Button
            {
                Text = "取消",
                Location = new Point(280, 610),
                Size = new Size(80, 30),
                Parent = form
            };
            
            // 更新按钮点击事件
            btnUpdate.Click += (s, e) =>
            {
                if (chkSendKeyPress.Checked)
                {
                    // 键盘按键输入模式
                    if (string.IsNullOrEmpty(recordedKey))
                    {
                        MessageBox.Show("请先录制按键", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return;
                    }
                    
                    if (!int.TryParse(txtWaitTime.Text, out int waitMs) || waitMs < 0)
                    {
                        MessageBox.Show("等待时间必须是非负整数", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return;
                    }
                    
                    // 验证鼠标滚动参数
                    int scrollCount = 1;
                    int scrollWaitMs = 100;
                    if (chkMouseScroll.Checked)
                    {
                        if (!int.TryParse(txtScrollCount.Text, out scrollCount) || scrollCount < 1)
                        {
                            MessageBox.Show("滚动次数必须是正整数", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            return;
                        }
                        
                        if (!int.TryParse(txtScrollWait.Text, out scrollWaitMs) || scrollWaitMs < 0)
                        {
                            MessageBox.Show("滚动后延时必须是非负整数", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            return;
                        }
                    }
                    
                    jumpMole.SendKeyPress = true;
                    jumpMole.KeyPressDefinition = recordedKey;
                    jumpMole.KeyPressWaitMs = waitMs;
                    jumpMole.EnableMouseScroll = chkMouseScroll.Checked;
                    jumpMole.ScrollUp = comboScrollDirection.SelectedIndex == 0;
                    jumpMole.ScrollCount = scrollCount;
                    jumpMole.ScrollWaitMs = scrollWaitMs;
                    jumpMole.Name = $"⌨️ 按键: {recordedKey}";
                    
                    SaveMoles();
                    
                    var lstMoles = GetCurrentMoleListBox();
                    if (lstMoles != null)
                    {
                        lstMoles.Items[moleIndex] = jumpMole.Name;
                    }
                    
                    var logMsg = $"✅ 已更新按键步骤: {recordedKey} (等待 {waitMs}ms)";
                    if (chkMouseScroll.Checked)
                    {
                        var direction = comboScrollDirection.SelectedIndex == 0 ? "向上" : "向下";
                        logMsg += $" + 鼠标{direction}滚动{scrollCount}次 (延时 {scrollWaitMs}ms)";
                    }
                    AppendLog(logMsg, LogType.Success);
                    form.Close();
                }
                else
                {
                    // 跳转模式
                    if (comboGroup.SelectedIndex < 0)
                    {
                        MessageBox.Show("请选择跳转目标分组", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return;
                    }
                    
                    var targetGroupName = comboGroup.SelectedItem.ToString();
                    var stepIndex = comboStep.SelectedIndex - 1; // -1 表示从头开始
                    
                    jumpMole.SendKeyPress = false;
                    jumpMole.JumpTargetGroup = targetGroupName;
                    jumpMole.JumpTargetStep = stepIndex;
                    jumpMole.Name = stepIndex < 0 
                        ? $"🔗 跳转到 {targetGroupName}" 
                        : $"🔗 跳转到 {targetGroupName} (步骤 {stepIndex + 1})";
                    
                    SaveMoles();
                    
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
                // 确保恢复热键（防止重复，先检查）
                if (hotkeysUnregistered)
                {
                    RegisterGlobalHotKeys();
                    hotkeysUnregistered = false;
                }
                
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
            
            // 自动聚焦删除按钮
            btnDelete.Focus();
        }

        private void ShowConfigStepDialog(MoleItem? configMole, int moleIndex)
        {
            var currentGroup = GetCurrentMoleGroup();
            if (currentGroup == null)
                return;
            
            bool isEdit = configMole != null;
            
            // 创建对话框
            var form = new Form
            {
                Text = isEdit ? "编辑配置步骤" : "添加配置步骤",
                Size = new Size(500, 400),
                StartPosition = FormStartPosition.Manual,
                FormBorderStyle = FormBorderStyle.FixedDialog,
                MaximizeBox = false,
                MinimizeBox = false
            };
            
            // 设置对话框位置：左边与主窗口右边对齐
            form.Location = new Point(this.Right, this.Top + (this.Height - form.Height) / 2);
            
            int yPos = 20;
            
            // ===== 操作1: 切换配置 =====
            var grpConfig = new GroupBox
            {
                Text = "操作1: 切换配置",
                Location = new Point(20, yPos),
                Size = new Size(450, 120),
                Parent = form
            };
            
            var chkSwitchConfig = new CheckBox
            {
                Text = "启用切换配置",
                Location = new Point(10, 25),
                Size = new Size(150, 25),
                Checked = configMole?.SwitchConfig ?? false,
                Parent = grpConfig
            };
            
            var lblConfig = new Label
            {
                Text = "配置:",
                Location = new Point(10, 55),
                Size = new Size(60, 20),
                Parent = grpConfig
            };
            
            var cmbConfig = new ComboBox
            {
                Location = new Point(70, 52),
                Size = new Size(200, 25),
                DropDownStyle = ComboBoxStyle.DropDownList,
                Parent = grpConfig
            };
            
            // 加载配置列表
            var configsDir = _configManager.ConfigsDirectory;
            if (Directory.Exists(configsDir))
            {
                var configFiles = Directory.GetFiles(configsDir, "*.json");
                foreach (var configFile in configFiles)
                {
                    var fileName = Path.GetFileNameWithoutExtension(configFile);
                    cmbConfig.Items.Add(fileName);
                }
            }
            
            if (cmbConfig.Items.Count > 0)
            {
                if (isEdit && !string.IsNullOrEmpty(configMole.TargetConfigName))
                {
                    int idx = cmbConfig.Items.IndexOf(configMole.TargetConfigName);
                    cmbConfig.SelectedIndex = idx >= 0 ? idx : 0;
                }
                else
                {
                    cmbConfig.SelectedIndex = 0;
                }
            }
            
            var lblConfigWait = new Label
            {
                Text = "等待:",
                Location = new Point(280, 55),
                Size = new Size(50, 20),
                Parent = grpConfig
            };
            
            var txtConfigWait = new TextBox
            {
                Location = new Point(330, 52),
                Size = new Size(60, 25),
                Text = (configMole?.ConfigSwitchWaitMs ?? 100).ToString(),
                Parent = grpConfig
            };
            
            var lblConfigMs = new Label
            {
                Text = "ms",
                Location = new Point(395, 55),
                Size = new Size(20, 20),
                Parent = grpConfig
            };
            
            yPos += 130;
            
            // ===== 操作2: 切换填充内容 =====
            var grpText = new GroupBox
            {
                Text = "操作2: 切换填充内容",
                Location = new Point(20, yPos),
                Size = new Size(450, 120),
                Parent = form
            };
            
            var chkSwitchText = new CheckBox
            {
                Text = "启用切换填充内容",
                Location = new Point(10, 25),
                Size = new Size(180, 25),
                Checked = configMole?.SwitchTextContent ?? false,
                Parent = grpText
            };
            
            var lblText = new Label
            {
                Text = "内容:",
                Location = new Point(10, 55),
                Size = new Size(60, 20),
                Parent = grpText
            };
            
            var cmbText = new ComboBox
            {
                Location = new Point(70, 52),
                Size = new Size(200, 25),
                DropDownStyle = ComboBoxStyle.DropDownList,
                Parent = grpText
            };
            
            // 加载文本内容列表
            foreach (var savedText in _config.SavedTexts)
            {
                cmbText.Items.Add(savedText.Name);
            }
            
            if (cmbText.Items.Count > 0)
            {
                if (isEdit && !string.IsNullOrEmpty(configMole.TargetTextName))
                {
                    int idx = cmbText.Items.IndexOf(configMole.TargetTextName);
                    cmbText.SelectedIndex = idx >= 0 ? idx : 0;
                }
                else
                {
                    cmbText.SelectedIndex = 0;
                }
            }
            
            var lblTextWait = new Label
            {
                Text = "等待:",
                Location = new Point(280, 55),
                Size = new Size(50, 20),
                Parent = grpText
            };
            
            var txtTextWait = new TextBox
            {
                Location = new Point(330, 52),
                Size = new Size(60, 25),
                Text = (configMole?.TextSwitchWaitMs ?? 100).ToString(),
                Parent = grpText
            };
            
            var lblTextMs = new Label
            {
                Text = "ms",
                Location = new Point(395, 55),
                Size = new Size(20, 20),
                Parent = grpText
            };
            
            yPos += 130;
            
            // 提示信息
            var lblHint = new Label
            {
                Text = "执行顺序: 配置切换 → 内容切换",
                Location = new Point(20, yPos),
                Size = new Size(450, 20),
                ForeColor = Color.Gray,
                Parent = form
            };
            
            yPos += 30;
            
            // 按钮
            var btnSave = new Button
            {
                Text = isEdit ? "保存" : "添加",
                Location = new Point(290, yPos),
                Size = new Size(80, 30),
                Parent = form
            };
            
            var btnCancel = new Button
            {
                Text = "取消",
                Location = new Point(380, yPos),
                Size = new Size(80, 30),
                Parent = form
            };
            
            // 如果是编辑模式，添加删除按钮
            Button? btnDelete = null;
            if (isEdit)
            {
                btnDelete = new Button
                {
                    Text = "删除",
                    Location = new Point(20, yPos),
                    Size = new Size(80, 30),
                    Parent = form
                };
                
                btnDelete.Click += (s, e) =>
                {
                    var result = MessageBox.Show(
                        $"确定要删除配置步骤吗？",
                        "确认删除",
                        MessageBoxButtons.YesNo,
                        MessageBoxIcon.Question);
                    
                    if (result == DialogResult.Yes)
                    {
                        currentGroup.Moles.RemoveAt(moleIndex);
                        SaveMoles();
                        RefreshCurrentMoleList();
                        AppendLog($"✅ 已删除配置步骤", LogType.Success);
                        form.Close();
                    }
                };
            }
            
            // 保存按钮
            btnSave.Click += (s, e) =>
            {
                if (!chkSwitchConfig.Checked && !chkSwitchText.Checked)
                {
                    MessageBox.Show("请至少选择一个操作", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
                
                if (chkSwitchConfig.Checked && cmbConfig.SelectedIndex < 0)
                {
                    MessageBox.Show("请选择目标配置", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
                
                if (chkSwitchText.Checked && cmbText.SelectedIndex < 0)
                {
                    MessageBox.Show("请选择目标填充内容", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
                
                if (!int.TryParse(txtConfigWait.Text, out int configWait) || configWait < 0)
                {
                    MessageBox.Show("配置切换等待时间必须是非负整数", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
                
                if (!int.TryParse(txtTextWait.Text, out int textWait) || textWait < 0)
                {
                    MessageBox.Show("内容切换等待时间必须是非负整数", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
                
                // 创建或更新配置步骤
                MoleItem stepMole;
                if (isEdit)
                {
                    stepMole = configMole!;
                }
                else
                {
                    stepMole = new MoleItem
                    {
                        IsConfigStep = true,
                        IsEnabled = true
                    };
                }
                
                stepMole.SwitchConfig = chkSwitchConfig.Checked;
                stepMole.TargetConfigName = cmbConfig.SelectedIndex >= 0 ? cmbConfig.Items[cmbConfig.SelectedIndex].ToString() ?? "" : "";
                stepMole.ConfigSwitchWaitMs = configWait;
                stepMole.SwitchTextContent = chkSwitchText.Checked;
                stepMole.TargetTextName = cmbText.SelectedIndex >= 0 ? cmbText.Items[cmbText.SelectedIndex].ToString() ?? "" : "";
                stepMole.TextSwitchWaitMs = textWait;
                
                // 生成步骤名称
                if (stepMole.SwitchConfig && stepMole.SwitchTextContent)
                {
                    stepMole.Name = $"⚙️ 配置: {stepMole.TargetConfigName} → 内容: {stepMole.TargetTextName}";
                }
                else if (stepMole.SwitchConfig)
                {
                    stepMole.Name = $"⚙️ 配置: {stepMole.TargetConfigName}";
                }
                else if (stepMole.SwitchTextContent)
                {
                    stepMole.Name = $"⚙️ 内容: {stepMole.TargetTextName}";
                }
                else
                {
                    stepMole.Name = "⚙️ 配置步骤 (未设置)";
                }
                
                if (!isEdit)
                {
                    currentGroup.Moles.Add(stepMole);
                }
                
                SaveMoles();
                RefreshCurrentMoleList();
                
                var action = isEdit ? "已更新" : "已添加";
                AppendLog($"✅ {action}配置步骤: {stepMole.Name}", LogType.Success);
                form.Close();
            };
            
            btnCancel.Click += (s, e) => form.Close();
            
            // 应用主题
            _themeManager.ApplyTheme(form);
            
            // 窗口关闭时清除引用
            form.FormClosed += (s, e) =>
            {
                if (_currentEditDialog == form)
                {
                    _currentEditDialog = null;
                }
            };
            
            // 显示对话框
            _currentEditDialog = form;
            form.Show();
            
            // 如果是编辑模式，自动聚焦删除按钮
            if (isEdit && btnDelete != null)
            {
                btnDelete.Focus();
            }
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
            
            // 等待超时后返回上一步复选框
            var chkReturnToPreviousOnTimeout = new CheckBox
            {
                Text = "等待超时后返回上一个步骤",
                Location = new Point(20, dialog.Height - 170),
                Size = new Size(200, 25),
                Checked = mole.ReturnToPreviousOnTimeout,
                Parent = dialog
            };
            
            // 超时时间标签
            var lblTimeoutLabel = new Label
            {
                Text = "超时时间:",
                Location = new Point(230, dialog.Height - 167),
                Size = new Size(70, 20),
                Parent = dialog
            };
            
            // 超时时间输入框
            var txtTimeout = new TextBox
            {
                Text = mole.TimeoutMs.ToString(),
                Location = new Point(300, dialog.Height - 170),
                Size = new Size(60, 25),
                Parent = dialog
            };
            
            // 超时时间单位标签
            var lblTimeoutUnit = new Label
            {
                Text = "ms",
                Location = new Point(365, dialog.Height - 167),
                Size = new Size(30, 20),
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
                
                // 验证超时时间
                if (!int.TryParse(txtTimeout.Text, out int timeoutMs))
                {
                    MessageBox.Show("请输入有效的超时时间数字", "输入错误", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
                
                if (timeoutMs < 0)
                {
                    MessageBox.Show("超时时间不能为负数", "输入错误", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
                
                // 保存所有设置
                mole.SimilarityThreshold = threshold;
                mole.ClickUntilDisappear = chkClickUntilDisappear.Checked;
                mole.WaitUntilAppear = chkWaitUntilAppear.Checked;
                mole.JumpToPreviousOnFail = chkJumpToPreviousOnFail.Checked;
                mole.ReturnToPreviousOnTimeout = chkReturnToPreviousOnTimeout.Checked;
                mole.TimeoutMs = timeoutMs;
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
                    
                    // 检查并处理 ImagePath
                    bool needsNewPath = false;
                    string oldPath = mole.ImagePath;
                    
                    // 检查路径是否为空或无效
                    if (string.IsNullOrWhiteSpace(mole.ImagePath))
                    {
                        needsNewPath = true;
                        AppendLog("⚠️ 图片路径为空，将生成新路径", LogType.Warning);
                    }
                    else if (!Path.IsPathRooted(mole.ImagePath))
                    {
                        // 相对路径，需要生成新路径
                        needsNewPath = true;
                        AppendLog($"⚠️ 检测到相对路径: {mole.ImagePath}，将生成新路径", LogType.Warning);
                    }
                    else
                    {
                        // 检查父目录是否存在
                        var parentDir = Path.GetDirectoryName(mole.ImagePath);
                        if (string.IsNullOrEmpty(parentDir) || !Directory.Exists(parentDir))
                        {
                            needsNewPath = true;
                            AppendLog($"⚠️ 父目录不存在: {parentDir}，将生成新路径", LogType.Warning);
                        }
                    }
                    
                    // 如果需要新路径，生成标准路径
                    if (needsNewPath)
                    {
                        var fileName = $"mole_{DateTime.Now:yyyyMMdd_HHmmss}.png";
                        mole.ImagePath = Path.Combine(_molesDirectory, fileName);
                        AppendLog($"✅ 已生成新路径: {mole.ImagePath}", LogType.Info);
                    }
                    else
                    {
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
                    }
                    
                    // 保存新截图
                    try
                    {
                        croppedImage.Save(mole.ImagePath, System.Drawing.Imaging.ImageFormat.Png);
                        croppedImage.Dispose();
                        
                        SaveMoles();
                        RefreshCurrentMoleList();
                        AppendLog($"✅ 已更新地鼠 \"{mole.Name}\" 的截图", LogType.Success);
                    }
                    catch (Exception ex)
                    {
                        AppendLog($"❌ 保存截图失败: {ex.Message}", LogType.Error);
                        MessageBox.Show($"保存截图失败: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        croppedImage.Dispose();
                    }
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

        // ==================== 加载设置相关方法 ====================
        
        private void ChkAutoLoadGroups_CheckedChanged(object? sender, EventArgs e)
        {
            _config.AutoLoadMoleGroups = chkAutoLoadGroups.Checked;
            SaveCurrentConfig();
            AppendLog($"✅ 自动显示已{(chkAutoLoadGroups.Checked ? "启用" : "禁用")}", LogType.Info);
        }

        private void BtnLoadSelectedGroups_Click(object? sender, EventArgs e)
        {
            LoadSelectedMoleGroups();
            // 切换到打地鼠标签页
            tabMain.SelectedTab = tabPageMole;
        }

        private void ChkSelectAllGroups_CheckedChanged(object? sender, EventArgs e)
        {
            if (lstMoleGroupsSelection.Items.Count == 0)
                return;

            // 避免递归触发
            lstMoleGroupsSelection.ItemCheck -= LstMoleGroupsSelection_ItemCheck;
            
            for (int i = 0; i < lstMoleGroupsSelection.Items.Count; i++)
            {
                lstMoleGroupsSelection.SetItemChecked(i, chkSelectAllGroups.Checked);
            }
            
            lstMoleGroupsSelection.ItemCheck += LstMoleGroupsSelection_ItemCheck;
            
            // 保存选择
            SaveMoleGroupSelection();
        }

        private void LstMoleGroupsSelection_ItemCheck(object? sender, ItemCheckEventArgs e)
        {
            // 延迟保存，因为此时 CheckedItems 还未更新
            BeginInvoke(new Action(() =>
            {
                SaveMoleGroupSelection();
            }));
        }

        private void SaveMoleGroupSelection()
        {
            _config.SelectedMoleGroups.Clear();
            foreach (int index in lstMoleGroupsSelection.CheckedIndices)
            {
                if (index < _moleGroups.Count)
                {
                    _config.SelectedMoleGroups.Add(_moleGroups[index].Name);
                }
            }
            SaveCurrentConfig();
        }

        private void LoadMoleGroupsSelection()
        {
            if (lstMoleGroupsSelection == null)
                return;
            
            // 临时移除事件处理器，避免在初始化时触发 BeginInvoke
            lstMoleGroupsSelection.ItemCheck -= LstMoleGroupsSelection_ItemCheck;
            
            lstMoleGroupsSelection.Items.Clear();
            
            foreach (var group in _moleGroups)
            {
                lstMoleGroupsSelection.Items.Add(group.Name);
            }

            // 恢复选择状态
            if (_config.SelectedMoleGroups.Count > 0)
            {
                for (int i = 0; i < _moleGroups.Count; i++)
                {
                    if (_config.SelectedMoleGroups.Contains(_moleGroups[i].Name))
                    {
                        lstMoleGroupsSelection.SetItemChecked(i, true);
                    }
                }
            }

            // 更新自动加载复选框状态
            if (chkAutoLoadGroups != null)
            {
                chkAutoLoadGroups.Checked = _config.AutoLoadMoleGroups;
            }
            
            // 重新添加事件处理器
            lstMoleGroupsSelection.ItemCheck += LstMoleGroupsSelection_ItemCheck;
        }

        private void LoadSelectedMoleGroups()
        {
            // 清空现有标签页
            tabMoleGroups.TabPages.Clear();

            // 获取选中的分组索引
            var selectedIndices = lstMoleGroupsSelection.CheckedIndices.Cast<int>().ToList();
            
            if (selectedIndices.Count == 0)
            {
                AppendLog("⚠️ 请至少选择一个分组", LogType.Warning);
                return;
            }

            // 只为选中的分组创建标签页
            foreach (int index in selectedIndices)
            {
                if (index < _moleGroups.Count)
                {
                    CreateMoleGroupTab(_moleGroups[index], index);
                }
            }

            // 选中第一个标签页
            if (tabMoleGroups.TabPages.Count > 0)
            {
                tabMoleGroups.SelectedIndex = 0;
                _currentMoleGroupIndex = selectedIndices[0];
            }

            AppendLog($"✅ 已显示 {selectedIndices.Count} 个分组", LogType.Success);
        }

        private void BtnExportGroups_Click(object? sender, EventArgs e)
        {
            // 获取选中的分组索引
            var selectedIndices = lstMoleGroupsSelection.CheckedIndices.Cast<int>().ToList();
            
            if (selectedIndices.Count == 0)
            {
                MessageBox.Show("请至少选择一个分组进行导出", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            try
            {
                // 获取程序所在目录
                var programDir = AppDomain.CurrentDomain.BaseDirectory;
                var exportDir = Path.Combine(programDir, "导出");
                
                // 确保导出目录存在
                if (!Directory.Exists(exportDir))
                {
                    Directory.CreateDirectory(exportDir);
                }

                // 为每个选中的分组创建导出文件
                foreach (int index in selectedIndices)
                {
                    if (index < _moleGroups.Count)
                    {
                        var group = _moleGroups[index];
                        ExportMoleGroup(group, exportDir);
                    }
                }

                AppendLog($"✅ 已导出 {selectedIndices.Count} 个分组到: {exportDir}", LogType.Success);
                
                // 弹窗提示导出成功
                MessageBox.Show($"导出成功！\n\n已导出 {selectedIndices.Count} 个分组到:\n{exportDir}", "导出成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"导出失败: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                AppendLog($"❌ 导出失败: {ex.Message}", LogType.Error);
            }
        }

        private void ExportMoleGroup(MoleGroup group, string exportDir)
        {
            // 创建分组专属文件夹
            var groupDir = Path.Combine(exportDir, group.Name);
            if (!Directory.Exists(groupDir))
            {
                Directory.CreateDirectory(groupDir);
            }

            // 创建图片文件夹
            var imagesDir = Path.Combine(groupDir, "images");
            if (!Directory.Exists(imagesDir))
            {
                Directory.CreateDirectory(imagesDir);
            }

            // 复制图片文件并更新路径
            var exportGroup = new MoleGroup
            {
                Name = group.Name,
                Moles = new List<MoleItem>()
            };

            foreach (var mole in group.Moles)
            {
                var exportMole = new MoleItem
                {
                    Name = mole.Name,
                    ImagePath = mole.ImagePath,
                    IsEnabled = mole.IsEnabled,
                    CreatedTime = mole.CreatedTime,
                    IsIdleClick = mole.IsIdleClick,
                    IdleClickPosition = mole.IdleClickPosition,
                    SimilarityThreshold = mole.SimilarityThreshold,
                    IsJump = mole.IsJump,
                    JumpTargetGroup = mole.JumpTargetGroup,
                    JumpTargetStep = mole.JumpTargetStep,
                    ClickUntilDisappear = mole.ClickUntilDisappear,
                    WaitUntilAppear = mole.WaitUntilAppear,
                    JumpToPreviousOnFail = mole.JumpToPreviousOnFail,
                    StopHunting = mole.StopHunting,
                    WaitAfterClick = mole.WaitAfterClick,
                    WaitAfterClickMs = mole.WaitAfterClickMs,
                    SendKeyPress = mole.SendKeyPress,
                    KeyPressDefinition = mole.KeyPressDefinition,
                    KeyPressWaitMs = mole.KeyPressWaitMs,
                    EnableMouseScroll = mole.EnableMouseScroll,
                    ScrollUp = mole.ScrollUp,
                    ScrollCount = mole.ScrollCount,
                    ScrollWaitMs = mole.ScrollWaitMs,
                    IsConfigStep = mole.IsConfigStep,
                    SwitchConfig = mole.SwitchConfig,
                    TargetConfigName = mole.TargetConfigName,
                    ConfigSwitchWaitMs = mole.ConfigSwitchWaitMs,
                    SwitchTextContent = mole.SwitchTextContent,
                    TargetTextName = mole.TargetTextName,
                    TextSwitchWaitMs = mole.TextSwitchWaitMs
                };

                // 如果有图片文件，复制到导出目录
                if (!string.IsNullOrEmpty(mole.ImagePath) && File.Exists(mole.ImagePath) && !mole.IsIdleClick && !mole.IsJump && !mole.IsConfigStep)
                {
                    var fileName = Path.GetFileName(mole.ImagePath);
                    var destPath = Path.Combine(imagesDir, fileName);
                    File.Copy(mole.ImagePath, destPath, true);
                    
                    // 更新为相对路径
                    exportMole.ImagePath = Path.Combine("images", fileName);
                }
                else
                {
                    exportMole.ImagePath = "";
                }

                exportGroup.Moles.Add(exportMole);
            }

            // 保存分组配置
            var configPath = Path.Combine(groupDir, "group_config.json");
            var json = Newtonsoft.Json.JsonConvert.SerializeObject(exportGroup, Newtonsoft.Json.Formatting.Indented);
            File.WriteAllText(configPath, json);
        }

        private void BtnImportGroups_Click(object? sender, EventArgs e)
        {
            try
            {
                // 获取程序所在目录
                var programDir = AppDomain.CurrentDomain.BaseDirectory;
                var exportDir = Path.Combine(programDir, "导出");
                
                // 确保导出目录存在
                if (!Directory.Exists(exportDir))
                {
                    Directory.CreateDirectory(exportDir);
                }

                // 使用 FolderBrowserDialog 让用户选择文件夹
                using (var fbd = new FolderBrowserDialog())
                {
                    fbd.Description = "选择要导入的分组文件夹（可以选择多个分组的父文件夹）";
                    fbd.SelectedPath = exportDir;
                    fbd.ShowNewFolderButton = false;

                    if (fbd.ShowDialog() != DialogResult.OK)
                        return;

                    var selectedPath = fbd.SelectedPath;
                    var importedGroups = new List<string>();
                    var renamedGroups = new List<(string oldName, string newName)>();

                    // 查找所有包含 group_config.json 的子文件夹
                    var configFiles = Directory.GetFiles(selectedPath, "group_config.json", SearchOption.AllDirectories);

                    if (configFiles.Length == 0)
                    {
                        MessageBox.Show("所选文件夹中没有找到分组配置文件", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        return;
                    }

                    // 导入所有找到的分组
                    foreach (var configPath in configFiles)
                    {
                        var result = ImportMoleGroup(configPath);
                        if (result.success)
                        {
                            importedGroups.Add(result.groupName);
                            if (result.renamed)
                            {
                                renamedGroups.Add((result.originalName, result.groupName));
                            }
                        }
                    }

                    if (importedGroups.Count > 0)
                    {
                        // 保存配置
                        SaveMoles();
                        
                        // 刷新显示设置界面
                        LoadMoleGroupsSelection();
                        
                        // 自动选中新导入的分组
                        for (int i = 0; i < _moleGroups.Count; i++)
                        {
                            if (importedGroups.Contains(_moleGroups[i].Name))
                            {
                                lstMoleGroupsSelection.SetItemChecked(i, true);
                            }
                        }
                        
                        // 自动加载并切换到打地鼠界面
                        LoadSelectedMoleGroups();
                        tabMain.SelectedTab = tabPageMole;
                        
                        AppendLog($"✅ 已导入 {importedGroups.Count} 个分组", LogType.Success);
                        
                        // 只有在有重命名的分组时才提示用户
                        if (renamedGroups.Count > 0)
                        {
                            var message = "以下分组因名称冲突已自动重命名：\n\n";
                            foreach (var (oldName, newName) in renamedGroups)
                            {
                                message += $"{oldName} → {newName}\n";
                            }
                            MessageBox.Show(message, "导入完成", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"导入失败: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                AppendLog($"❌ 导入失败: {ex.Message}", LogType.Error);
            }
        }

        private (bool success, string groupName, string originalName, bool renamed) ImportMoleGroup(string configPath)
        {
            try
            {
                // 读取配置文件
                var json = File.ReadAllText(configPath);
                var importGroup = Newtonsoft.Json.JsonConvert.DeserializeObject<MoleGroup>(json);
                
                if (importGroup == null)
                {
                    return (false, "", "", false);
                }

                var originalName = importGroup.Name;
                var groupDir = Path.GetDirectoryName(configPath);
                var imagesDir = Path.Combine(groupDir!, "images");

                // 检查名称冲突并自动重命名
                var finalName = importGroup.Name;
                var renamed = false;
                var counter = 2;
                
                while (_moleGroups.Any(g => g.Name == finalName))
                {
                    finalName = $"{importGroup.Name}_{counter}";
                    counter++;
                    renamed = true;
                }

                importGroup.Name = finalName;

                // 处理图片文件
                foreach (var mole in importGroup.Moles)
                {
                    if (!string.IsNullOrEmpty(mole.ImagePath) && !mole.IsIdleClick && !mole.IsJump && !mole.IsConfigStep)
                    {
                        var sourceImagePath = Path.Combine(groupDir!, mole.ImagePath);
                        
                        if (File.Exists(sourceImagePath))
                        {
                            // 生成唯一的文件名
                            var fileName = Path.GetFileName(sourceImagePath);
                            var destPath = Path.Combine(_molesDirectory, $"{finalName}_{fileName}");
                            
                            // 如果文件已存在，添加时间戳
                            if (File.Exists(destPath))
                            {
                                var nameWithoutExt = Path.GetFileNameWithoutExtension(fileName);
                                var ext = Path.GetExtension(fileName);
                                destPath = Path.Combine(_molesDirectory, $"{finalName}_{nameWithoutExt}_{DateTime.Now:yyyyMMddHHmmss}{ext}");
                            }
                            
                            File.Copy(sourceImagePath, destPath, true);
                            mole.ImagePath = destPath;
                        }
                        else
                        {
                            mole.ImagePath = "";
                        }
                    }
                }

                // 添加到分组列表
                _moleGroups.Add(importGroup);

                return (true, finalName, originalName, renamed);
            }
            catch (Exception ex)
            {
                AppendLog($"⚠️ 导入分组失败: {ex.Message}", LogType.Warning);
                return (false, "", "", false);
            }
        }

        // ==================== 加载设置相关方法结束 ====================

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
