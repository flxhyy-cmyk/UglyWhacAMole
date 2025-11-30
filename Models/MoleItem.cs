using System;
using System.Collections.Generic;
using System.Drawing;

namespace WindowInspector.Models
{
    /// <summary>
    /// 地鼠项（截图目标）
    /// </summary>
    public class MoleItem
    {
        public string Name { get; set; } = "";
        public string ImagePath { get; set; } = "";
        public bool IsEnabled { get; set; } = true;
        public DateTime CreatedTime { get; set; } = DateTime.Now;
        public bool IsIdleClick { get; set; } = false; // 是否为空击位置
        public Point? IdleClickPosition { get; set; } = null; // 空击坐标
        public double SimilarityThreshold { get; set; } = 0.85; // 匹配阈值（0-1），默认0.85
        public bool IsJump { get; set; } = false; // 是否为跳转步骤
        public string JumpTargetGroup { get; set; } = ""; // 跳转目标分组名称
        public int JumpTargetStep { get; set; } = -1; // 跳转目标步骤索引（-1表示从头开始）
        public bool ClickUntilDisappear { get; set; } = false; // 持续点击直到消失
        public bool WaitUntilAppear { get; set; } = false; // 持续等待直到出现
        public bool JumpToPreviousOnFail { get; set; } = false; // 识别失败时跳转到上一个步骤
        public bool StopHunting { get; set; } = false; // 执行到此步骤时停止打地鼠
        public bool WaitAfterClick { get; set; } = false; // 点击后等待
        public int WaitAfterClickMs { get; set; } = 1000; // 点击后等待的毫秒数
        public bool SendKeyPress { get; set; } = false; // 发送键盘按键输入（跳转步骤的替代选项）
        public string KeyPressDefinition { get; set; } = ""; // 按键定义（如 "Ctrl+C", "Enter", "F1"）
        public int KeyPressWaitMs { get; set; } = 100; // 按键输入后等待的毫秒数
        public bool EnableMouseScroll { get; set; } = false; // 是否启用鼠标滚动（在按键输入之后执行）
        public bool ScrollUp { get; set; } = true; // true=向上滚动, false=向下滚动
        public int ScrollCount { get; set; } = 1; // 滚动次数
        public int ScrollWaitMs { get; set; } = 100; // 滚动后等待时间（毫秒）
        public bool IsConfigStep { get; set; } = false; // 是否为配置步骤
        public bool SwitchConfig { get; set; } = false; // 是否切换配置
        public string TargetConfigName { get; set; } = ""; // 目标配置名称
        public int ConfigSwitchWaitMs { get; set; } = 100; // 配置切换后等待时间（毫秒）
        public bool SwitchTextContent { get; set; } = false; // 是否切换填充内容
        public string TargetTextName { get; set; } = ""; // 目标文本名称
        public int TextSwitchWaitMs { get; set; } = 100; // 内容切换后等待时间（毫秒）
        public bool ReturnToPreviousOnTimeout { get; set; } = false; // 等待超时后返回上一个步骤
        public int TimeoutMs { get; set; } = 5000; // 超时时间（毫秒），默认5000ms
    }

    /// <summary>
    /// 地鼠列表组（标签页）
    /// </summary>
    public class MoleGroup
    {
        public string Name { get; set; } = "默认";
        public List<MoleItem> Moles { get; set; } = new();
    }
}
