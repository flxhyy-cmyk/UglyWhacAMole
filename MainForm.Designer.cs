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
        private CheckBox chkMoleEnabled;
        private Button btnCaptureMole;
        private Button btnSetIdleClick;
        private Button btnAddConfigStep;
        private Button btnBatchSelect;
        private Button btnAddJump;
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
            this.FormBorderStyle = FormBorderStyle.FixedSingle;
            this.MaximizeBox = false;
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
            tabMoleGroups = new TabControl
            {
                Location = new Point(0, 130),
                Size = new Size(tabPageMole.ClientSize.Width, tabPageMole.ClientSize.Height - 130),
                Padding = new Point(0, 0),
                Margin = new Padding(0),
                SizeMode = TabSizeMode.Fixed,
                ItemSize = new Size(80, 25),
                Appearance = TabAppearance.Buttons,
                Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right,
                Parent = tabPageMole
            };
            tabMoleGroups.SelectedIndexChanged += TabMoleGroups_SelectedIndexChanged;
            tabMoleGroups.MouseDoubleClick += TabMoleGroups_MouseDoubleClick;
        }


    }
}
