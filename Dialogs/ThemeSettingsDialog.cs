using System;
using System.Drawing;
using System.Windows.Forms;
using WindowInspector.Models;
using WindowInspector.Utils;

namespace WindowInspector.Dialogs
{
    /// <summary>
    /// 主题设置对话框
    /// </summary>
    public class ThemeSettingsDialog : Form
    {
        private ComboBox cmbTheme;
        private Button btnOk;
        private Button btnCancel;
        private Label lblTheme;
        private Label lblDescription;
        private ThemeMode _selectedTheme;
        private readonly ThemeManager _themeManager;

        public ThemeMode SelectedTheme => _selectedTheme;

        public ThemeSettingsDialog(ThemeManager themeManager)
        {
            _themeManager = themeManager;
            _selectedTheme = themeManager.Settings.Mode;
            InitializeComponent();
            LoadCurrentTheme();
            
            // 应用当前主题到对话框
            _themeManager.ApplyTheme(this);
        }

        private void InitializeComponent()
        {
            Text = "主题设置";
            Size = new Size(400, 220);
            StartPosition = FormStartPosition.CenterParent;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;

            // 标题标签
            lblTheme = new Label
            {
                Text = "选择主题:",
                Location = new Point(20, 20),
                Size = new Size(350, 25),
                Font = new Font(Font.FontFamily, 10, FontStyle.Bold),
                Parent = this
            };

            // 主题选择下拉框
            cmbTheme = new ComboBox
            {
                Location = new Point(20, 50),
                Size = new Size(350, 30),
                DropDownStyle = ComboBoxStyle.DropDownList,
                Parent = this
            };

            cmbTheme.Items.Add("🌞 浅色主题");
            cmbTheme.Items.Add("🌙 深色主题");
            cmbTheme.Items.Add("🔄 随系统切换");

            cmbTheme.SelectedIndexChanged += CmbTheme_SelectedIndexChanged;

            // 描述标签
            lblDescription = new Label
            {
                Name = "lblDescription",
                Text = GetThemeDescription(_selectedTheme),
                Location = new Point(20, 90),
                Size = new Size(350, 40),
                ForeColor = Color.Gray,
                Parent = this
            };

            // 确定按钮
            btnOk = new Button
            {
                Text = "确定",
                Location = new Point(200, 140),
                Size = new Size(80, 30),
                DialogResult = DialogResult.OK,
                Parent = this
            };

            // 取消按钮
            btnCancel = new Button
            {
                Text = "取消",
                Location = new Point(290, 140),
                Size = new Size(80, 30),
                DialogResult = DialogResult.Cancel,
                Parent = this
            };

            AcceptButton = btnOk;
            CancelButton = btnCancel;
        }

        private void LoadCurrentTheme()
        {
            cmbTheme.SelectedIndex = (int)_selectedTheme;
        }

        private void CmbTheme_SelectedIndexChanged(object? sender, EventArgs e)
        {
            _selectedTheme = (ThemeMode)cmbTheme.SelectedIndex;
            lblDescription.Text = GetThemeDescription(_selectedTheme);
        }

        private string GetThemeDescription(ThemeMode mode)
        {
            return mode switch
            {
                ThemeMode.Light => "使用浅色主题,适合明亮环境使用",
                ThemeMode.Dark => "使用深色主题,减少眼睛疲劳",
                ThemeMode.System => "自动跟随操作系统主题设置",
                _ => ""
            };
        }
    }
}
