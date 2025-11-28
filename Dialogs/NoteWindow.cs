using System;
using System.Drawing;
using System.Windows.Forms;
using WindowInspector.Models;
using WindowInspector.Utils;

namespace WindowInspector.Dialogs
{
    /// <summary>
    /// 便签窗口 - 支持主题切换
    /// </summary>
    public class NoteWindow : Form
    {
        private TextBox txtNote;
        private ComboBox cmbTheme;
        private Label lblTheme;
        private Button btnSave;
        private Button btnClear;
        private readonly ThemeManager _themeManager;
        private readonly ConfigManager _configManager;
        private readonly string _noteFilePath;

        public NoteWindow()
        {
            _configManager = new ConfigManager();
            _themeManager = new ThemeManager(_configManager);
            _noteFilePath = System.IO.Path.Combine(_configManager.ProgramDirectory, "note.txt");
            
            InitializeComponent();
            LoadNote();
            
            // 应用当前主题
            _themeManager.ApplyTheme(this);
        }

        private void InitializeComponent()
        {
            Text = "便签";
            Size = new Size(500, 400);
            StartPosition = FormStartPosition.CenterScreen;
            FormBorderStyle = FormBorderStyle.Sizable;
            MinimumSize = new Size(400, 300);

            // 主题选择标签
            lblTheme = new Label
            {
                Text = "主题:",
                Location = new Point(10, 15),
                Size = new Size(45, 20),
                Parent = this
            };

            // 主题选择下拉框
            cmbTheme = new ComboBox
            {
                Location = new Point(55, 12),
                Size = new Size(150, 25),
                DropDownStyle = ComboBoxStyle.DropDownList,
                Parent = this
            };

            cmbTheme.Items.Add("🌞 浅色主题");
            cmbTheme.Items.Add("🌙 深色主题");
            cmbTheme.Items.Add("🔄 随系统切换");

            cmbTheme.SelectedIndex = (int)_themeManager.Settings.Mode;
            cmbTheme.SelectedIndexChanged += CmbTheme_SelectedIndexChanged;

            // 保存按钮
            btnSave = new Button
            {
                Text = "保存",
                Location = new Point(220, 10),
                Size = new Size(80, 28),
                Parent = this
            };
            btnSave.Click += BtnSave_Click;

            // 清空按钮
            btnClear = new Button
            {
                Text = "清空",
                Location = new Point(310, 10),
                Size = new Size(80, 28),
                Parent = this
            };
            btnClear.Click += BtnClear_Click;

            // 便签内容文本框
            txtNote = new TextBox
            {
                Location = new Point(10, 45),
                Size = new Size(this.ClientSize.Width - 20, this.ClientSize.Height - 55),
                Multiline = true,
                ScrollBars = ScrollBars.Vertical,
                Font = new Font("Microsoft YaHei UI", 10),
                Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right,
                Parent = this
            };

            // 窗口大小改变时调整文本框大小
            Resize += (s, e) =>
            {
                txtNote.Size = new Size(ClientSize.Width - 20, ClientSize.Height - 55);
            };

            // 窗口关闭时自动保存
            FormClosing += (s, e) =>
            {
                SaveNote();
            };
        }

        private void CmbTheme_SelectedIndexChanged(object? sender, EventArgs e)
        {
            var selectedTheme = (ThemeMode)cmbTheme.SelectedIndex;
            _themeManager.ChangeTheme(selectedTheme);
            _themeManager.ApplyTheme(this);
        }

        private void BtnSave_Click(object? sender, EventArgs e)
        {
            SaveNote();
            MessageBox.Show("便签已保存!", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void BtnClear_Click(object? sender, EventArgs e)
        {
            var result = MessageBox.Show(
                "确定要清空便签内容吗?",
                "确认",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question);

            if (result == DialogResult.Yes)
            {
                txtNote.Clear();
                SaveNote();
            }
        }

        private void SaveNote()
        {
            try
            {
                System.IO.File.WriteAllText(_noteFilePath, txtNote.Text);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"保存失败: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void LoadNote()
        {
            try
            {
                if (System.IO.File.Exists(_noteFilePath))
                {
                    txtNote.Text = System.IO.File.ReadAllText(_noteFilePath);
                }
            }
            catch { }
        }
    }
}
