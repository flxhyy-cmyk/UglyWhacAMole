using System;
using System.Drawing;
using System.Windows.Forms;
using WindowInspector.Models;
using Microsoft.Win32;

namespace WindowInspector.Utils
{
    /// <summary>
    /// 主题管理器
    /// </summary>
    public class ThemeManager
    {
        private ThemeSettings _settings;
        private readonly ConfigManager _configManager;
        
        public ThemeManager(ConfigManager configManager)
        {
            _configManager = configManager;
            _settings = LoadThemeSettings() ?? new ThemeSettings();
        }

        /// <summary>
        /// 获取当前主题设置
        /// </summary>
        public ThemeSettings Settings => _settings;

        /// <summary>
        /// 保存主题设置
        /// </summary>
        public void SaveThemeSettings()
        {
            try
            {
                var json = Newtonsoft.Json.JsonConvert.SerializeObject(_settings, Newtonsoft.Json.Formatting.Indented);
                System.IO.File.WriteAllText(
                    System.IO.Path.Combine(_configManager.ProgramDirectory, "theme_settings.json"),
                    json);
            }
            catch { }
        }

        /// <summary>
        /// 加载主题设置
        /// </summary>
        private ThemeSettings? LoadThemeSettings()
        {
            try
            {
                var filePath = System.IO.Path.Combine(_configManager.ProgramDirectory, "theme_settings.json");
                if (!System.IO.File.Exists(filePath))
                    return null;

                var json = System.IO.File.ReadAllText(filePath);
                return Newtonsoft.Json.JsonConvert.DeserializeObject<ThemeSettings>(json);
            }
            catch
            {
                return null;
            }
        }

        /// <summary>
        /// 获取实际应用的主题(处理系统模式)
        /// </summary>
        public ThemeMode GetEffectiveTheme()
        {
            if (_settings.Mode == ThemeMode.System)
            {
                return IsSystemDarkMode() ? ThemeMode.Dark : ThemeMode.Light;
            }
            return _settings.Mode;
        }

        /// <summary>
        /// 判断系统是否为深色模式
        /// </summary>
        private bool IsSystemDarkMode()
        {
            try
            {
                using (var key = Registry.CurrentUser.OpenSubKey(@"Software\Microsoft\Windows\CurrentVersion\Themes\Personalize"))
                {
                    var value = key?.GetValue("AppsUseLightTheme");
                    return value is int i && i == 0;
                }
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// 应用主题到窗体
        /// </summary>
        public void ApplyTheme(Form form)
        {
            var effectiveTheme = GetEffectiveTheme();
            IThemeColors colors = effectiveTheme == ThemeMode.Dark ? DarkThemeColors.Instance : LightThemeColors.Instance;

            // 设置窗体背景色
            form.BackColor = colors.Background;
            form.ForeColor = colors.Foreground;

            // 递归应用到所有控件
            ApplyThemeToControl(form, colors);

            // 再次确保窗体背景色正确（防止被子控件覆盖）
            form.BackColor = colors.Background;
            form.ForeColor = colors.Foreground;
        }

        /// <summary>
        /// 递归应用主题到控件
        /// </summary>
        private void ApplyThemeToControl(Control control, IThemeColors colors)
        {
            var effectiveTheme = GetEffectiveTheme();
            
            // 根据控件类型应用不同样式
            switch (control)
            {
                case Button btn:
                    btn.BackColor = colors.ButtonBackground;
                    btn.ForeColor = colors.ButtonForeground;
                    btn.FlatStyle = FlatStyle.Flat;
                    btn.FlatAppearance.BorderColor = colors.BorderColor;
                    btn.FlatAppearance.BorderSize = 1;
                    // 鼠标悬停时稍微变亮
                    btn.FlatAppearance.MouseOverBackColor = Color.FromArgb(
                        Math.Min(255, ((Color)colors.ButtonBackground).R + 20),
                        Math.Min(255, ((Color)colors.ButtonBackground).G + 20),
                        Math.Min(255, ((Color)colors.ButtonBackground).B + 20));
                    break;

                case TextBox txt:
                    txt.BackColor = colors.InputBackground;
                    txt.ForeColor = colors.InputForeground;
                    txt.BorderStyle = BorderStyle.FixedSingle;
                    break;

                case ComboBox cmb:
                    cmb.BackColor = colors.InputBackground;
                    cmb.ForeColor = colors.InputForeground;
                    cmb.FlatStyle = FlatStyle.Flat;
                    // 设置下拉按钮颜色（通过重绘实现）
                    if (cmb.DrawMode != DrawMode.OwnerDrawFixed && cmb.Name != "cmbSavedTexts")
                    {
                        cmb.DrawMode = DrawMode.OwnerDrawFixed;
                        cmb.DrawItem -= ComboBox_DrawItem;
                        cmb.DrawItem += ComboBox_DrawItem;
                    }
                    break;

                case RichTextBox rtb:
                    // 保持日志窗口的原有深色主题
                    if (rtb.Name != "rtbLog")
                    {
                        rtb.BackColor = colors.InputBackground;
                        rtb.ForeColor = colors.InputForeground;
                    }
                    else
                    {
                        // 日志窗口根据主题调整
                        if (effectiveTheme == ThemeMode.Dark)
                        {
                            rtb.BackColor = colors.LogBackground;
                            rtb.ForeColor = colors.LogForeground;
                        }
                        else
                        {
                            // 浅色模式：浅色背景，深色文字
                            rtb.BackColor = Color.FromArgb(250, 250, 250);  // 几乎白色
                            rtb.ForeColor = Color.FromArgb(30, 30, 30);     // 深色文字
                        }
                    }
                    break;

                case CheckedListBox clb:
                    clb.BackColor = colors.InputBackground;
                    clb.ForeColor = colors.InputForeground;
                    
                    // 如果控件有 "CustomDraw" 标记，跳过主题绘制接管和 BorderStyle 修改
                    if (clb.Tag?.ToString() == "CustomDraw")
                    {
                        // 只设置颜色，不修改 BorderStyle 和绘制模式
                        break;
                    }
                    
                    clb.BorderStyle = BorderStyle.FixedSingle;
                    // 启用自定义绘制以支持深色模式
                    clb.DrawMode = DrawMode.OwnerDrawFixed;
                    clb.DrawItem -= CheckedListBox_DrawItem; // 先移除避免重复
                    clb.DrawItem += CheckedListBox_DrawItem;
                    break;

                case ListBox lb:
                    lb.BackColor = colors.InputBackground;
                    lb.ForeColor = colors.InputForeground;
                    lb.BorderStyle = BorderStyle.FixedSingle;
                    break;

                case CheckBox chk:
                    chk.BackColor = colors.Background;
                    chk.ForeColor = colors.Foreground;
                    break;

                case RadioButton rb:
                    rb.BackColor = colors.Background;
                    rb.ForeColor = colors.Foreground;
                    break;

                case Label lbl:
                    // 保持某些标签的特殊颜色(如空击位置标签)
                    if (lbl.Name == "lblIdleClickPos")
                    {
                        // 空击位置标签保持特殊颜色
                        if (lbl.ForeColor != Color.Gray && lbl.ForeColor != Color.Green)
                        {
                            lbl.ForeColor = colors.Foreground;
                        }
                        lbl.BackColor = Color.Transparent;
                    }
                    else if (lbl.Name == "lblDescription")
                    {
                        // 描述标签在浅色模式下使用深灰色，深色模式下使用浅灰色
                        lbl.BackColor = Color.Transparent;
                        lbl.ForeColor = effectiveTheme == ThemeMode.Dark 
                            ? Color.FromArgb(180, 180, 180)  // 深色模式：浅灰色
                            : Color.FromArgb(100, 100, 100); // 浅色模式：深灰色
                    }
                    else
                    {
                        lbl.BackColor = Color.Transparent;
                        // 标签文字使用更亮的颜色
                        lbl.ForeColor = effectiveTheme == ThemeMode.Dark
                            ? Color.FromArgb(220, 220, 220)  // 深色模式：浅色文字
                            : Color.FromArgb(50, 50, 50);    // 浅色模式：深色文字
                    }
                    break;

                case GroupBox grp:
                    grp.BackColor = colors.Background;
                    grp.ForeColor = Color.FromArgb(220, 220, 220);  // GroupBox标题更亮
                    // 自定义绘制GroupBox边框
                    grp.Paint -= GroupBox_Paint;
                    grp.Paint += GroupBox_Paint;
                    break;

                case TabControl tab:
                    tab.BackColor = colors.Background;
                    tab.ForeColor = colors.Foreground;
                    // 启用自定义绘制以支持深色模式
                    tab.DrawMode = TabDrawMode.OwnerDrawFixed;
                    tab.DrawItem -= TabControl_DrawItem; // 先移除避免重复
                    tab.DrawItem += TabControl_DrawItem;
                    // 添加Paint事件来覆盖边框
                    tab.Paint -= TabControl_Paint;
                    tab.Paint += TabControl_Paint;
                    // 扩展TabControl的区域以覆盖边框
                    if (tab.Parent != null && effectiveTheme == ThemeMode.Dark)
                    {
                        tab.Location = new Point(tab.Location.X - 2, tab.Location.Y - 2);
                        tab.Size = new Size(tab.Width + 4, tab.Height + 4);
                    }
                    break;

                case TabPage page:
                    page.BackColor = colors.Background;
                    page.ForeColor = colors.Foreground;
                    break;

                case Panel panel:
                    // 保持Caps Lock指示器的功能性颜色
                    if (panel.Name != "pnlCapsIndicator")
                    {
                        panel.BackColor = colors.Background;
                        panel.ForeColor = colors.Foreground;
                    }
                    break;

                case Form form:
                    // 确保窗体背景色正确应用
                    form.BackColor = colors.Background;
                    form.ForeColor = colors.Foreground;
                    break;

                case VScrollBar vScrollBar:
                    // VScrollBar在Windows Forms中难以自定义，但我们可以设置基本颜色
                    vScrollBar.BackColor = colors.InputBackground;
                    vScrollBar.ForeColor = colors.Foreground;
                    break;

                case HScrollBar hScrollBar:
                    hScrollBar.BackColor = colors.InputBackground;
                    hScrollBar.ForeColor = colors.Foreground;
                    break;

                case ScrollBar scrollBar:
                    // ScrollBar在Windows Forms中难以自定义，但我们可以设置基本颜色
                    scrollBar.BackColor = colors.InputBackground;
                    scrollBar.ForeColor = colors.Foreground;
                    break;
            }

            // 递归处理子控件
            foreach (Control child in control.Controls)
            {
                ApplyThemeToControl(child, colors);
            }
        }

        /// <summary>
        /// 更改主题
        /// </summary>
        public void ChangeTheme(ThemeMode mode)
        {
            _settings.Mode = mode;
            SaveThemeSettings();
        }

        /// <summary>
        /// GroupBox自定义绘制边框
        /// </summary>
        private void GroupBox_Paint(object? sender, PaintEventArgs e)
        {
            if (sender is not GroupBox groupBox) return;

            var effectiveTheme = GetEffectiveTheme();
            if (effectiveTheme != ThemeMode.Dark) return;

            IThemeColors colors = DarkThemeColors.Instance;

            // 清除默认边框
            e.Graphics.Clear(groupBox.BackColor);

            // 测量标题文字大小
            SizeF textSize = e.Graphics.MeasureString(groupBox.Text, groupBox.Font);

            // 绘制边框（避开标题文字）
            using (var pen = new Pen(colors.BorderColor, 1))
            {
                int textWidth = (int)textSize.Width + 10;
                int textHeight = (int)textSize.Height;

                // 上边框（分两段，避开文字）
                e.Graphics.DrawLine(pen, 0, textHeight / 2, 8, textHeight / 2);
                e.Graphics.DrawLine(pen, 8 + textWidth, textHeight / 2, groupBox.Width - 1, textHeight / 2);

                // 左、右、下边框
                e.Graphics.DrawLine(pen, 0, textHeight / 2, 0, groupBox.Height - 1);
                e.Graphics.DrawLine(pen, groupBox.Width - 1, textHeight / 2, groupBox.Width - 1, groupBox.Height - 1);
                e.Graphics.DrawLine(pen, 0, groupBox.Height - 1, groupBox.Width - 1, groupBox.Height - 1);
            }

            // 绘制标题文字
            using (var brush = new SolidBrush(groupBox.ForeColor))
            {
                e.Graphics.DrawString(groupBox.Text, groupBox.Font, brush, 10, 0);
            }
        }

        /// <summary>
        /// ComboBox自定义绘制
        /// </summary>
        private void ComboBox_DrawItem(object? sender, DrawItemEventArgs e)
        {
            if (sender is not ComboBox comboBox) return;
            if (e.Index < 0) return;

            var effectiveTheme = GetEffectiveTheme();
            IThemeColors colors = effectiveTheme == ThemeMode.Dark ? DarkThemeColors.Instance : LightThemeColors.Instance;

            e.DrawBackground();

            Color textColor = effectiveTheme == ThemeMode.Dark 
                ? Color.FromArgb(240, 240, 240) 
                : SystemColors.WindowText;

            using (var brush = new SolidBrush(textColor))
            {
                e.Graphics.DrawString(
                    comboBox.Items[e.Index].ToString(),
                    e.Font ?? comboBox.Font,
                    brush,
                    e.Bounds);
            }

            e.DrawFocusRectangle();
        }

        /// <summary>
        /// CheckedListBox自定义绘制事件处理
        /// </summary>
        private void CheckedListBox_DrawItem(object? sender, DrawItemEventArgs e)
        {
            if (sender is not CheckedListBox checkedListBox) return;
            if (e.Index < 0 || e.Index >= checkedListBox.Items.Count) return;

            var effectiveTheme = GetEffectiveTheme();
            IThemeColors colors = effectiveTheme == ThemeMode.Dark ? DarkThemeColors.Instance : LightThemeColors.Instance;

            // 绘制背景
            Color backColor;
            Color textColor;

            if ((e.State & DrawItemState.Selected) == DrawItemState.Selected)
            {
                // 选中项 - 使用更明显的高亮色
                backColor = effectiveTheme == ThemeMode.Dark 
                    ? Color.FromArgb(70, 70, 70) 
                    : SystemColors.Highlight;
                textColor = effectiveTheme == ThemeMode.Dark 
                    ? Color.FromArgb(255, 255, 255)  // 纯白色文字
                    : SystemColors.HighlightText;
            }
            else
            {
                // 未选中项
                backColor = colors.InputBackground;
                textColor = colors.InputForeground;
            }

            using (var backBrush = new SolidBrush(backColor))
            {
                e.Graphics.FillRectangle(backBrush, e.Bounds);
            }

            // 绘制复选框
            Rectangle checkBoxRect = new Rectangle(e.Bounds.Left + 2, e.Bounds.Top + 2, 16, 16);
            bool isChecked = checkedListBox.GetItemChecked(e.Index);

            // 绘制复选框背景
            Color checkBoxBackColor = effectiveTheme == ThemeMode.Dark 
                ? Color.FromArgb(50, 50, 50) 
                : SystemColors.Window;
            using (var checkBoxBrush = new SolidBrush(checkBoxBackColor))
            {
                e.Graphics.FillRectangle(checkBoxBrush, checkBoxRect);
            }

            // 绘制复选框边框
            using (var checkBoxPen = new Pen(colors.BorderColor))
            {
                e.Graphics.DrawRectangle(checkBoxPen, checkBoxRect);
            }

            // 如果选中，绘制勾选标记
            if (isChecked)
            {
                using (var checkPen = new Pen(effectiveTheme == ThemeMode.Dark 
                    ? Color.FromArgb(100, 200, 100) 
                    : Color.Green, 2))
                {
                    e.Graphics.DrawLine(checkPen, 
                        checkBoxRect.Left + 3, checkBoxRect.Top + 8,
                        checkBoxRect.Left + 6, checkBoxRect.Top + 11);
                    e.Graphics.DrawLine(checkPen, 
                        checkBoxRect.Left + 6, checkBoxRect.Top + 11,
                        checkBoxRect.Left + 13, checkBoxRect.Top + 4);
                }
            }

            // 绘制文本
            Rectangle textRect = new Rectangle(
                e.Bounds.Left + 22, 
                e.Bounds.Top, 
                e.Bounds.Width - 22, 
                e.Bounds.Height);

            using (var textBrush = new SolidBrush(textColor))
            {
                StringFormat stringFormat = new StringFormat
                {
                    LineAlignment = StringAlignment.Center,
                    Trimming = StringTrimming.EllipsisCharacter
                };
                e.Graphics.DrawString(
                    checkedListBox.Items[e.Index].ToString(), 
                    e.Font ?? checkedListBox.Font, 
                    textBrush, 
                    textRect, 
                    stringFormat);
            }

            // 绘制焦点矩形
            if ((e.State & DrawItemState.Focus) == DrawItemState.Focus)
            {
                ControlPaint.DrawFocusRectangle(e.Graphics, e.Bounds);
            }
        }

        /// <summary>
        /// TabControl Paint事件 - 覆盖白色边框
        /// </summary>
        private void TabControl_Paint(object? sender, PaintEventArgs e)
        {
            if (sender is not TabControl tabControl) return;

            var effectiveTheme = GetEffectiveTheme();
            if (effectiveTheme != ThemeMode.Dark) return;

            IThemeColors colors = DarkThemeColors.Instance;

            // 获取标签页区域的高度
            int tabHeight = tabControl.ItemSize.Height + 4;

            // 覆盖TabControl的白色边框区域（更大的覆盖范围）
            using (var brush = new SolidBrush(colors.Background))
            {
                // 覆盖标签页周围的所有白色边框
                // 顶部和标签区域
                Rectangle topArea = new Rectangle(0, 0, tabControl.Width, tabHeight);
                e.Graphics.FillRectangle(brush, topArea);

                // 左侧边框（从标签下方开始）
                Rectangle leftBorder = new Rectangle(0, tabHeight, 4, tabControl.Height - tabHeight);
                e.Graphics.FillRectangle(brush, leftBorder);

                // 右侧边框（从标签下方开始）
                Rectangle rightBorder = new Rectangle(tabControl.Width - 4, tabHeight, 4, tabControl.Height - tabHeight);
                e.Graphics.FillRectangle(brush, rightBorder);

                // 底部边框
                Rectangle bottomBorder = new Rectangle(0, tabControl.Height - 4, tabControl.Width, 4);
                e.Graphics.FillRectangle(brush, bottomBorder);
            }

            // 重新绘制标签页按钮（因为被覆盖了）
            for (int i = 0; i < tabControl.TabCount; i++)
            {
                Rectangle tabBounds = tabControl.GetTabRect(i);
                DrawTabButton(e.Graphics, tabControl, i, tabBounds);
            }
        }

        private void DrawTabButton(Graphics g, TabControl tabControl, int index, Rectangle bounds)
        {
            var effectiveTheme = GetEffectiveTheme();
            IThemeColors colors = effectiveTheme == ThemeMode.Dark ? DarkThemeColors.Instance : LightThemeColors.Instance;

            bool isSelected = tabControl.SelectedIndex == index;
            Color backColor = isSelected 
                ? (effectiveTheme == ThemeMode.Dark ? Color.FromArgb(45, 45, 45) : SystemColors.ControlLightLight)
                : (effectiveTheme == ThemeMode.Dark ? Color.FromArgb(35, 35, 35) : SystemColors.Control);
            Color textColor = isSelected
                ? (effectiveTheme == ThemeMode.Dark ? Color.FromArgb(255, 255, 255) : SystemColors.ControlText)
                : (effectiveTheme == ThemeMode.Dark ? Color.FromArgb(180, 180, 180) : SystemColors.ControlText);

            using (var backBrush = new SolidBrush(backColor))
            {
                g.FillRectangle(backBrush, bounds);
            }

            if (isSelected)
            {
                using (var linePen = new Pen(backColor, 3))
                {
                    g.DrawLine(linePen, bounds.Left, bounds.Bottom, bounds.Right, bounds.Bottom);
                }
            }

            using (var textBrush = new SolidBrush(textColor))
            {
                StringFormat stringFormat = new StringFormat
                {
                    Alignment = StringAlignment.Center,
                    LineAlignment = StringAlignment.Center
                };
                g.DrawString(tabControl.TabPages[index].Text, tabControl.Font, textBrush, bounds, stringFormat);
            }
        }

        /// <summary>
        /// TabControl自定义绘制事件处理
        /// </summary>
        private void TabControl_DrawItem(object? sender, DrawItemEventArgs e)
        {
            if (sender is not TabControl tabControl) return;

            var effectiveTheme = GetEffectiveTheme();
            IThemeColors colors = effectiveTheme == ThemeMode.Dark ? DarkThemeColors.Instance : LightThemeColors.Instance;

            Graphics g = e.Graphics;
            TabPage tabPage = tabControl.TabPages[e.Index];
            Rectangle tabBounds = tabControl.GetTabRect(e.Index);

            // 在第一个标签时，先填充整个标签条区域的背景
            if (e.Index == 0)
            {
                Rectangle tabStripRect = new Rectangle(0, 0, tabControl.Width, tabBounds.Height + 4);
                Color stripBackColor = effectiveTheme == ThemeMode.Dark 
                    ? Color.FromArgb(35, 35, 35)  // 标签条背景色
                    : SystemColors.Control;
                using (var stripBrush = new SolidBrush(stripBackColor))
                {
                    g.FillRectangle(stripBrush, tabStripRect);
                }
            }

            // 绘制背景
            Color backColor;
            Color textColor;

            if (e.State == DrawItemState.Selected)
            {
                // 选中的标签页 - 使用更亮的背景色，与内容区一致
                backColor = effectiveTheme == ThemeMode.Dark 
                    ? Color.FromArgb(45, 45, 45)  // 与主背景色一致
                    : SystemColors.ControlLightLight;
                textColor = effectiveTheme == ThemeMode.Dark 
                    ? Color.FromArgb(255, 255, 255)  // 纯白色文字
                    : SystemColors.ControlText;
            }
            else
            {
                // 未选中的标签页 - 使用更暗的背景色以区分
                backColor = effectiveTheme == ThemeMode.Dark 
                    ? Color.FromArgb(35, 35, 35)  // 与标签条背景一致
                    : SystemColors.Control;
                textColor = effectiveTheme == ThemeMode.Dark 
                    ? Color.FromArgb(180, 180, 180)  // 灰色文字
                    : SystemColors.ControlText;
            }

            // 扩展绘制区域，消除缝隙
            Rectangle expandedBounds = new Rectangle(
                tabBounds.X, 
                tabBounds.Y, 
                tabBounds.Width, 
                tabBounds.Height + 2);

            using (var backBrush = new SolidBrush(backColor))
            {
                g.FillRectangle(backBrush, expandedBounds);
            }

            // 只在选中标签的底部绘制一条线，与内容区连接
            if (e.State == DrawItemState.Selected)
            {
                using (var linePen = new Pen(backColor, 3))
                {
                    g.DrawLine(linePen, 
                        tabBounds.Left, tabBounds.Bottom,
                        tabBounds.Right, tabBounds.Bottom);
                }
            }

            // 绘制文本 - "打地鼠"标签不需要特殊颜色，使用默认颜色即可
            // 移除了特殊颜色处理，让所有标签使用统一的配色方案

            using (var textBrush = new SolidBrush(textColor))
            {
                StringFormat stringFormat = new StringFormat
                {
                    Alignment = StringAlignment.Center,
                    LineAlignment = StringAlignment.Center
                };
                g.DrawString(tabPage.Text, tabControl.Font, textBrush, tabBounds, stringFormat);
            }
        }
    }

    /// <summary>
    /// 主题颜色接口
    /// </summary>
    public interface IThemeColors
    {
        Color Background { get; }
        Color Foreground { get; }
        Color ButtonBackground { get; }
        Color ButtonForeground { get; }
        Color InputBackground { get; }
        Color InputForeground { get; }
        Color BorderColor { get; }
        Color LogBackground { get; }
        Color LogForeground { get; }
    }

    /// <summary>
    /// 浅色主题颜色
    /// </summary>
    public class LightThemeColors : IThemeColors
    {
        public static LightThemeColors Instance { get; } = new LightThemeColors();

        public Color Background => SystemColors.Control;
        public Color Foreground => SystemColors.ControlText;
        public Color ButtonBackground => SystemColors.Control;
        public Color ButtonForeground => SystemColors.ControlText;
        public Color InputBackground => SystemColors.Window;
        public Color InputForeground => SystemColors.WindowText;
        public Color BorderColor => SystemColors.ControlDark;
        public Color LogBackground => Color.FromArgb(43, 43, 43);
        public Color LogForeground => Color.White;
    }

    /// <summary>
    /// 深色主题颜色
    /// </summary>
    public class DarkThemeColors : IThemeColors
    {
        public static DarkThemeColors Instance { get; } = new DarkThemeColors();

        public Color Background => Color.FromArgb(45, 45, 45);  // 主背景色
        public Color Foreground => Color.FromArgb(240, 240, 240);  // 主文字色
        public Color ButtonBackground => Color.FromArgb(70, 70, 70);  // 按钮更亮，增强对比
        public Color ButtonForeground => Color.FromArgb(255, 255, 255);  // 按钮文字纯白
        public Color InputBackground => Color.FromArgb(30, 30, 30);  // 输入框更暗以区分
        public Color InputForeground => Color.FromArgb(240, 240, 240);
        public Color BorderColor => Color.FromArgb(100, 100, 100);  // 边框更明显
        public Color LogBackground => Color.FromArgb(35, 35, 35);  // 日志区域与整体更协调
        public Color LogForeground => Color.FromArgb(240, 240, 240);
    }
}
