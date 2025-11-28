using System;

namespace WindowInspector.Models
{
    /// <summary>
    /// 主题设置
    /// </summary>
    public class ThemeSettings
    {
        /// <summary>
        /// 主题模式
        /// </summary>
        public ThemeMode Mode { get; set; } = ThemeMode.Light;
    }

    /// <summary>
    /// 主题模式枚举
    /// </summary>
    public enum ThemeMode
    {
        /// <summary>
        /// 浅色主题
        /// </summary>
        Light,
        
        /// <summary>
        /// 深色主题
        /// </summary>
        Dark,
        
        /// <summary>
        /// 跟随系统
        /// </summary>
        System
    }
}
