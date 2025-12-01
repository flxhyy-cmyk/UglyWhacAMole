using System;
using System.Windows.Forms;

namespace WindowInspector.Models
{
    /// <summary>
    /// 快捷键设置
    /// </summary>
    public class HotkeySettings
    {
        /// <summary>
        /// F2功能快捷键（填充文本）
        /// </summary>
        public Keys F2Key { get; set; } = Keys.F2;

        /// <summary>
        /// F3功能快捷键（打地鼠开关）
        /// </summary>
        public Keys F3Key { get; set; } = Keys.F3;

        /// <summary>
        /// F4功能快捷键（截图创建地鼠）
        /// </summary>
        public Keys F4Key { get; set; } = Keys.F4;

        /// <summary>
        /// F6功能快捷键（添加空击位置）
        /// </summary>
        public Keys F6Key { get; set; } = Keys.F6;

        /// <summary>
        /// 配置文本定义快捷键（默认无）
        /// </summary>
        public Keys? ConfigTextKey { get; set; } = null;

        /// <summary>
        /// 批量启用/禁用快捷键（默认无）
        /// </summary>
        public Keys? BatchSelectKey { get; set; } = null;

        /// <summary>
        /// 添加跳转/键鼠快捷键（默认无）
        /// </summary>
        public Keys? AddJumpKey { get; set; } = null;
    }
}
