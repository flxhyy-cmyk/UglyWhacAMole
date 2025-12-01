using System;
using System.Drawing;
using System.Windows.Forms;

namespace WindowInspector
{
    partial class MainForm
    {
        private System.ComponentModel.IContainer components = null;
        private Button btnSelectWindow;
        private Button btnRecordInput;
        private Button btnSaveText;
        private Button btnLoadExcel;
        private Button btnOpenExcel;
        private Button btnFillText;
        private Button btnExportExcel;
        private Button btnConfigOps;
        private ComboBox cmbSavedTexts;
        private ComboBox cmbCellGroups;
        private TextBox txtInputCount;
        private RichTextBox rtbLog;
        private Label lblInputCount;
        private Panel pnlCapsIndicator;
        private GroupBox grpOperations;
        private GroupBox grpStatus;
        
        // 打地鼠相关控件
        private CheckBox chkContinuousClick;
        private CheckBox chkFullScreenMatch;
        private TabControl tabMain;
        private TabPage tabPageMain;
        private TabPage tabPageMole;
        private TabPage tabPageLoadSettings;
        private TabPage tabPageHotkeySettings;
        private CheckBox chkMoleEnabled;
        private CheckBox chkAutoLoadGroups;
        private Button btnLoadSelectedGroups;
        private Button btnExportGroups;
        private Button btnImportGroups;
        private CheckBox chkSelectAllGroups;
        private CheckedListBox lstMoleGroupsSelection;
        private Button btnCaptureMole;
        private Button btnSetIdleClick;
        private Button btnAddConfigStep;
        private Button btnBatchSelect;
        private Button btnAddJump;
        private Button btnMoveStep;
        private Label lblIdleClickPos;
        private CheckedListBox lstMoles;
        private TabControl tabMoleGroups;
        private Button btnAddMoleGroup;
        private Button btnRemoveMoleGroup;

        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            this.Text = "文本框位置记录工具";
            this.ClientSize = new Size(504, 520);  // 默认高度520
            this.StartPosition = FormStartPosition.CenterScreen;
            this.FormBorderStyle = FormBorderStyle.Sizable;  // 允许鼠标拖拽调整窗口大小
            this.MaximizeBox = true;
            this.Padding = new Padding(0);

            // 创建标签页控件
            tabMain = new TabControl
            {
                Location = new Point(0, 0),
                Size = new Size(this.ClientSize.Width, this.ClientSize.Height),
                Padding = new Point(0, 0),
                Margin = new Padding(0),
                SizeMode = TabSizeMode.Fixed,
                ItemSize = new Size(100, 30),
                Appearance = TabAppearance.Buttons,
                Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right,
                Parent = this
            };

            // 主功能标签页
            tabPageMain = new TabPage
            {
                Text = "文本填充",
                Padding = new Padding(0),
                Margin = new Padding(0),
                Parent = tabMain
            };

            // 打地鼠标签页
            tabPageMole = new TabPage
            {
                Text = "打地鼠",
                Padding = new Padding(0),
                Margin = new Padding(0),
                Parent = tabMain
            };

            // 显示设置标签页
            tabPageLoadSettings = new TabPage
            {
                Text = "显示设置",
                Padding = new Padding(0),
                Margin = new Padding(0),
                Parent = tabMain
            };

            // 快捷键设置标签页
            tabPageHotkeySettings = new TabPage
            {
                Text = "快捷键设置",
                Padding = new Padding(0),
                Margin = new Padding(0),
                Parent = tabMain
            };

            // 操作区（缩小高度）
            grpOperations = new GroupBox
            {
                Text = "操作区",
                Location = new Point(0, 0),
                Size = new Size(tabPageMain.ClientSize.Width, 170),
                Padding = new Padding(0),
                Margin = new Padding(0),
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right,
                Parent = tabPageMain
            };

            // 第一行按钮
            btnSelectWindow = new Button
            {
                Text = "1. 选择目标窗口",
                Location = new Point(0, 25),
                Size = new Size(140, 30),
                Parent = grpOperations
            };
            btnSelectWindow.Click += BtnSelectWindow_Click;

            btnRecordInput = new Button
            {
                Text = "2. 记录输入框位置",
                Location = new Point(150, 25),
                Size = new Size(160, 30),
                Enabled = false,
                Parent = grpOperations
            };
            btnRecordInput.Click += BtnRecordInput_Click;

            lblInputCount = new Label
            {
                Text = "数量:",
                Location = new Point(320, 30),
                Size = new Size(40, 20),
                Parent = grpOperations
            };

            txtInputCount = new TextBox
            {
                Text = "2",
                Location = new Point(365, 27),
                Size = new Size(40, 25),
                Parent = grpOperations
            };

            // 单元格组选择（Excel模式下显示，位置在第一行下方）
            cmbCellGroups = new ComboBox
            {
                Location = new Point(150, 60),
                Size = new Size(160, 25),
                DropDownStyle = ComboBoxStyle.DropDownList,
                Visible = false,
                Parent = grpOperations
            };

            // 第二行按钮（紧贴第一行，Y=65）
            btnSaveText = new Button
            {
                Text = "保存输入内容",
                Location = new Point(0, 65),
                Size = new Size(110, 30),
                Parent = grpOperations
            };
            btnSaveText.Click += BtnSaveText_Click;
            btnSaveText.MouseDown += BtnSaveText_MouseDown;

            btnLoadExcel = new Button
            {
                Text = "加载",
                Location = new Point(120, 65),
                Size = new Size(80, 30),
                Parent = grpOperations
            };
            btnLoadExcel.Click += BtnLoadExcel_Click;
            btnLoadExcel.MouseDown += BtnLoadExcel_MouseDown;

            btnOpenExcel = new Button
            {
                Text = "打开",
                Location = new Point(210, 65),
                Size = new Size(80, 30),
                Parent = grpOperations
            };
            btnOpenExcel.Click += BtnOpenExcel_Click;

            btnFillText = new Button
            {
                Text = "填充",
                Location = new Point(300, 65),
                Size = new Size(80, 30),
                Parent = grpOperations
            };
            btnFillText.Click += BtnFillText_Click;

            // 第三行（Y=105）
            cmbSavedTexts = new ComboBox
            {
                Location = new Point(0, 105),
                Size = new Size(200, 25),
                DropDownStyle = ComboBoxStyle.DropDownList,
                Parent = grpOperations
            };

            btnExportExcel = new Button
            {
                Text = "导出",
                Location = new Point(210, 105),
                Size = new Size(70, 40),
                Parent = grpOperations
            };
            btnExportExcel.Click += BtnExportExcel_Click;

            btnConfigOps = new Button
            {
                Text = "配置操作",
                Location = new Point(290, 105),
                Size = new Size(90, 40),
                Parent = grpOperations
            };
            btnConfigOps.MouseDown += BtnConfigOps_MouseDown;

            pnlCapsIndicator = new Panel
            {
                Name = "pnlCapsIndicator",
                Location = new Point(400, 105),
                Size = new Size(20, 20),
                BackColor = Color.Green,
                Parent = grpOperations
            };

            // 状态信息区（向上移动）
            grpStatus = new GroupBox
            {
                Text = "状态信息",
                Location = new Point(0, 170),
                Size = new Size(tabPageMain.ClientSize.Width, tabPageMain.ClientSize.Height - 170),
                Padding = new Padding(0),
                Margin = new Padding(0),
                Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right,
                Parent = tabPageMain
            };

            rtbLog = new RichTextBox
            {
                Name = "rtbLog",
                Location = new Point(0, 20),
                Size = new Size(grpStatus.ClientSize.Width, grpStatus.ClientSize.Height - 20),
                Padding = new Padding(0),
                Margin = new Padding(0),
                Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right,
                BackColor = Color.FromArgb(43, 43, 43),
                ForeColor = Color.White,
                ReadOnly = true,
                BorderStyle = BorderStyle.None,
                Parent = grpStatus
            };

            // 打地鼠界面
            InitializeMoleTab();
            
            // 显示设置界面
            InitializeLoadSettingsTab();
            
            // 快捷键设置界面
            InitializeHotkeySettingsTab();
        }

        private void InitializeMoleTab()
        {
            // 功能开关
            chkMoleEnabled = new CheckBox
            {
                Text = "启用打地鼠 (F3)",
                Location = new Point(10, 10),
                Size = new Size(150, 25),
                Parent = tabPageMole
            };
            chkMoleEnabled.CheckedChanged += ChkMoleEnabled_CheckedChanged;

            // 全图匹配复选框
            chkFullScreenMatch = new CheckBox
            {
                Text = "全图匹配",
                Location = new Point(10, 40),
                Size = new Size(90, 25),
                Checked = false,
                Parent = tabPageMole
            };
            chkFullScreenMatch.CheckedChanged += (s, e) => 
            {
                _moleHunter?.SetFullScreenMatch(chkFullScreenMatch.Checked);
            };

            // 持续点击开关
            chkContinuousClick = new CheckBox
            {
                Text = "持续点击直到消失",
                Location = new Point(110, 40),
                Size = new Size(150, 25),
                Checked = false,
                Parent = tabPageMole
            };
            chkContinuousClick.CheckedChanged += (s, e) => 
            {
                _moleHunter?.SetContinuousClick(chkContinuousClick.Checked);
            };

            // 截图按钮
            btnCaptureMole = new Button
            {
                Text = "截图创建地鼠 (F4)",
                Location = new Point(170, 7),
                Size = new Size(140, 30),
                Parent = tabPageMole
            };
            btnCaptureMole.Click += BtnCaptureMole_Click;

            // 空击按钮
            btnSetIdleClick = new Button
            {
                Text = "添加空击位置 (F6)",
                Location = new Point(320, 7),
                Size = new Size(140, 30),
                Parent = tabPageMole
            };
            btnSetIdleClick.Click += BtnSetIdleClick_Click;

            // 空击位置显示
            lblIdleClickPos = new Label
            {
                Name = "lblIdleClickPos",
                Text = "空击: 未设置",
                Location = new Point(320, 42),
                Size = new Size(140, 20),
                ForeColor = Color.Gray,
                Parent = tabPageMole
            };

            // 配置文本定义按钮（放在批量选择按钮前面）
            btnAddConfigStep = new Button
            {
                Text = "配置文本定义",
                Location = new Point(10, 65),
                Size = new Size(120, 30),
                Parent = tabPageMole
            };
            btnAddConfigStep.Click += BtnAddConfigStep_Click;
            
            // 批量选择按钮
            btnBatchSelect = new Button
            {
                Text = "批量启用/禁用",
                Location = new Point(140, 65),
                Size = new Size(120, 30),
                Parent = tabPageMole
            };
            btnBatchSelect.Click += BtnBatchSelect_Click;
            
            // 添加跳转/键鼠按钮
            btnAddJump = new Button
            {
                Text = "添加跳转/键鼠",
                Location = new Point(270, 65),
                Size = new Size(120, 30),
                Parent = tabPageMole
            };
            btnAddJump.Click += BtnAddJump_Click;

            // 步骤移动按钮
            btnMoveStep = new Button
            {
                Text = "↕",
                Location = new Point(400, 65),
                Size = new Size(40, 30),
                Font = new Font("Arial", 14, FontStyle.Bold),
                Parent = tabPageMole
            };
            btnMoveStep.Click += BtnMoveStep_Click;
            
            // 地鼠列表标签页控件
            var lblMoles = new Label
            {
                Text = "地鼠列表（右键删除）:",
                Location = new Point(10, 105),
                Size = new Size(150, 20),
                Parent = tabPageMole
            };

            // 添加/删除标签页按钮
            btnAddMoleGroup = new Button
            {
                Text = "+",
                Location = new Point(370, 102),
                Size = new Size(30, 25),
                Parent = tabPageMole
            };
            btnAddMoleGroup.Click += BtnAddMoleGroup_Click;

            btnRemoveMoleGroup = new Button
            {
                Text = "-",
                Location = new Point(410, 102),
                Size = new Size(30, 25),
                Parent = tabPageMole
            };
            btnRemoveMoleGroup.Click += BtnRemoveMoleGroup_Click;

            // 地鼠列表标签页
            // 注意：TabControl 的 ItemSize.Height (25px) 会占用内部空间，需要在计算时考虑
            // 底部保留1像素间距
            tabMoleGroups = new TabControl
            {
                Location = new Point(1, 133),
                Size = new Size(tabPageMole.ClientSize.Width - 2, Math.Max(tabPageMole.ClientSize.Height - 133 - 1, 100)),
                Padding = new Point(0, 0),
                Margin = new Padding(0),
                SizeMode = TabSizeMode.Fixed,
                ItemSize = new Size(80, 25),
                Appearance = TabAppearance.Buttons,
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right, // 移除 Bottom，手动控制高度
                Parent = tabPageMole
            };
            tabMoleGroups.SelectedIndexChanged += TabMoleGroups_SelectedIndexChanged;
            tabMoleGroups.MouseDoubleClick += TabMoleGroups_MouseDoubleClick;
            // 监听 Resize 事件，动态调整 CheckedListBox 的高度
            tabMoleGroups.Resize += TabMoleGroups_Resize;
            // 监听父容器 Resize 事件，手动调整 tabMoleGroups 高度以保持底部1像素间距
            tabPageMole.Resize += (s, e) =>
            {
                tabMoleGroups.Height = tabPageMole.ClientSize.Height - 130 - 1;
            };
        }

        private void InitializeLoadSettingsTab()
        {
            // 自动显示复选框
            chkAutoLoadGroups = new CheckBox
            {
                Text = "启用自动显示",
                Location = new Point(10, 10),
                Size = new Size(150, 25),
                Parent = tabPageLoadSettings
            };
            chkAutoLoadGroups.CheckedChanged += ChkAutoLoadGroups_CheckedChanged;

            // 显示按钮
            btnLoadSelectedGroups = new Button
            {
                Text = "显示所选分组",
                Location = new Point(170, 7),
                Size = new Size(120, 30),
                Parent = tabPageLoadSettings
            };
            btnLoadSelectedGroups.Click += BtnLoadSelectedGroups_Click;

            // 导出按钮
            btnExportGroups = new Button
            {
                Text = "导出",
                Location = new Point(300, 7),
                Size = new Size(70, 30),
                Parent = tabPageLoadSettings
            };
            btnExportGroups.Click += BtnExportGroups_Click;

            // 导入按钮
            btnImportGroups = new Button
            {
                Text = "导入",
                Location = new Point(380, 7),
                Size = new Size(70, 30),
                Parent = tabPageLoadSettings
            };
            btnImportGroups.Click += BtnImportGroups_Click;

            // 分组列表标签
            var lblGroups = new Label
            {
                Text = "选择要显示的分组:",
                Location = new Point(10, 45),
                Size = new Size(200, 20),
                Parent = tabPageLoadSettings
            };

            // 全选复选框
            chkSelectAllGroups = new CheckBox
            {
                Text = "全选/取消全选",
                Location = new Point(220, 43),
                Size = new Size(150, 25),
                Parent = tabPageLoadSettings
            };
            chkSelectAllGroups.CheckedChanged += ChkSelectAllGroups_CheckedChanged;

            // 分组选择列表
            lstMoleGroupsSelection = new CheckedListBox
            {
                Location = new Point(10, 70),
                Size = new Size(tabPageLoadSettings.ClientSize.Width - 20, tabPageLoadSettings.ClientSize.Height - 80),
                Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right,
                CheckOnClick = true,
                Parent = tabPageLoadSettings
            };
            lstMoleGroupsSelection.ItemCheck += LstMoleGroupsSelection_ItemCheck;
        }

        private TextBox txtHotkeyF2;
        private TextBox txtHotkeyF3;
        private TextBox txtHotkeyF4;
        private TextBox txtHotkeyF6;
        private TextBox txtHotkeyConfigText;
        private TextBox txtHotkeyBatchSelect;
        private TextBox txtHotkeyAddJump;

        private void InitializeHotkeySettingsTab()
        {
            // 说明标签
            var lblDescription = new Label
            {
                Text = "点击输入框后按下想要设置的快捷键，失去焦点时自动保存",
                Location = new Point(20, 20),
                Size = new Size(450, 25),
                Font = new Font(Font.FontFamily, 9, FontStyle.Regular),
                ForeColor = Color.Gray,
                Parent = tabPageHotkeySettings
            };

            // F2功能
            var lblF2 = new Label
            {
                Text = "填充文本:",
                Location = new Point(20, 60),
                Size = new Size(150, 25),
                Parent = tabPageHotkeySettings
            };

            txtHotkeyF2 = new TextBox
            {
                Name = "txtHotkeyF2",
                Location = new Point(180, 58),
                Size = new Size(250, 25),
                ReadOnly = true,
                Parent = tabPageHotkeySettings
            };
            txtHotkeyF2.KeyDown += TxtHotkeyF2_KeyDown;
            txtHotkeyF2.Enter += TxtHotkey_Enter;
            txtHotkeyF2.Leave += TxtHotkey_Leave;

            // F3功能
            var lblF3 = new Label
            {
                Text = "打地鼠开关:",
                Location = new Point(20, 100),
                Size = new Size(150, 25),
                Parent = tabPageHotkeySettings
            };

            txtHotkeyF3 = new TextBox
            {
                Name = "txtHotkeyF3",
                Location = new Point(180, 98),
                Size = new Size(250, 25),
                ReadOnly = true,
                Parent = tabPageHotkeySettings
            };
            txtHotkeyF3.KeyDown += TxtHotkeyF3_KeyDown;
            txtHotkeyF3.Enter += TxtHotkey_Enter;
            txtHotkeyF3.Leave += TxtHotkey_Leave;

            // F4功能
            var lblF4 = new Label
            {
                Text = "截图创建地鼠:",
                Location = new Point(20, 140),
                Size = new Size(150, 25),
                Parent = tabPageHotkeySettings
            };

            txtHotkeyF4 = new TextBox
            {
                Name = "txtHotkeyF4",
                Location = new Point(180, 138),
                Size = new Size(250, 25),
                ReadOnly = true,
                Parent = tabPageHotkeySettings
            };
            txtHotkeyF4.KeyDown += TxtHotkeyF4_KeyDown;
            txtHotkeyF4.Enter += TxtHotkey_Enter;
            txtHotkeyF4.Leave += TxtHotkey_Leave;

            // F6功能
            var lblF6 = new Label
            {
                Text = "添加空击位置:",
                Location = new Point(20, 180),
                Size = new Size(150, 25),
                Parent = tabPageHotkeySettings
            };

            txtHotkeyF6 = new TextBox
            {
                Name = "txtHotkeyF6",
                Location = new Point(180, 178),
                Size = new Size(250, 25),
                ReadOnly = true,
                Parent = tabPageHotkeySettings
            };
            txtHotkeyF6.KeyDown += TxtHotkeyF6_KeyDown;
            txtHotkeyF6.Enter += TxtHotkey_Enter;
            txtHotkeyF6.Leave += TxtHotkey_Leave;

            // 分隔线
            var lblSeparator = new Label
            {
                Text = "打地鼠功能快捷键:",
                Location = new Point(20, 220),
                Size = new Size(400, 20),
                Font = new Font(Font.FontFamily, 9, FontStyle.Bold),
                Parent = tabPageHotkeySettings
            };

            // 配置文本定义
            var lblConfigText = new Label
            {
                Text = "配置文本定义:",
                Location = new Point(20, 250),
                Size = new Size(150, 25),
                Parent = tabPageHotkeySettings
            };

            txtHotkeyConfigText = new TextBox
            {
                Name = "txtHotkeyConfigText",
                Location = new Point(180, 248),
                Size = new Size(250, 25),
                ReadOnly = true,
                Parent = tabPageHotkeySettings
            };
            txtHotkeyConfigText.KeyDown += TxtHotkeyConfigText_KeyDown;
            txtHotkeyConfigText.Enter += TxtHotkey_Enter;
            txtHotkeyConfigText.Leave += TxtHotkey_Leave;

            // 批量启用/禁用
            var lblBatchSelect = new Label
            {
                Text = "批量启用/禁用:",
                Location = new Point(20, 290),
                Size = new Size(150, 25),
                Parent = tabPageHotkeySettings
            };

            txtHotkeyBatchSelect = new TextBox
            {
                Name = "txtHotkeyBatchSelect",
                Location = new Point(180, 288),
                Size = new Size(250, 25),
                ReadOnly = true,
                Parent = tabPageHotkeySettings
            };
            txtHotkeyBatchSelect.KeyDown += TxtHotkeyBatchSelect_KeyDown;
            txtHotkeyBatchSelect.Enter += TxtHotkey_Enter;
            txtHotkeyBatchSelect.Leave += TxtHotkey_Leave;

            // 添加跳转/键鼠
            var lblAddJump = new Label
            {
                Text = "添加跳转/键鼠:",
                Location = new Point(20, 330),
                Size = new Size(150, 25),
                Parent = tabPageHotkeySettings
            };

            txtHotkeyAddJump = new TextBox
            {
                Name = "txtHotkeyAddJump",
                Location = new Point(180, 328),
                Size = new Size(250, 25),
                ReadOnly = true,
                Parent = tabPageHotkeySettings
            };
            txtHotkeyAddJump.KeyDown += TxtHotkeyAddJump_KeyDown;
            txtHotkeyAddJump.Enter += TxtHotkey_Enter;
            txtHotkeyAddJump.Leave += TxtHotkey_Leave;

            // 提示信息
            var lblHint = new Label
            {
                Text = "提示: 修改快捷键后需要重启程序才能生效",
                Location = new Point(20, 370),
                Size = new Size(400, 20),
                ForeColor = Color.Orange,
                Parent = tabPageHotkeySettings
            };

            // 重置按钮
            var btnResetHotkeys = new Button
            {
                Text = "恢复默认快捷键",
                Location = new Point(20, 400),
                Size = new Size(150, 30),
                Parent = tabPageHotkeySettings
            };
            btnResetHotkeys.Click += BtnResetHotkeys_Click;
        }

    }
}
