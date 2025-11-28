using System;
using System.Runtime.InteropServices;
using System.Text;
using System.Diagnostics;

namespace WindowInspector.Utils
{
    public static class WindowHelper
    {
        [DllImport("user32.dll")]
        public static extern IntPtr WindowFromPoint(POINT point);

        [DllImport("user32.dll")]
        public static extern IntPtr GetAncestor(IntPtr hwnd, uint gaFlags);

        [DllImport("user32.dll")]
        public static extern bool GetWindowRect(IntPtr hwnd, out RECT rect);

        [DllImport("user32.dll")]
        public static extern int GetClassName(IntPtr hWnd, StringBuilder lpClassName, int nMaxCount);

        [DllImport("user32.dll")]
        public static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString, int nMaxCount);

        [DllImport("user32.dll")]
        public static extern bool SetForegroundWindow(IntPtr hWnd);

        [DllImport("user32.dll")]
        public static extern bool SetCursorPos(int x, int y);

        [DllImport("user32.dll")]
        public static extern bool GetCursorPos(out POINT lpPoint);

        [DllImport("user32.dll")]
        public static extern void mouse_event(uint dwFlags, int dx, int dy, uint dwData, UIntPtr dwExtraInfo);

        [DllImport("user32.dll")]
        public static extern short GetAsyncKeyState(int vKey);

        [DllImport("user32.dll")]
        public static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, IntPtr lParam);

        [DllImport("user32.dll")]
        public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);

        [DllImport("imm32.dll")]
        public static extern IntPtr ImmGetContext(IntPtr hWnd);

        [DllImport("imm32.dll")]
        public static extern bool ImmReleaseContext(IntPtr hWnd, IntPtr hIMC);

        [DllImport("imm32.dll")]
        public static extern bool ImmSetOpenStatus(IntPtr hIMC, bool fOpen);

        [DllImport("imm32.dll")]
        public static extern bool ImmGetOpenStatus(IntPtr hIMC);

        [DllImport("imm32.dll")]
        public static extern IntPtr ImmAssociateContext(IntPtr hWnd, IntPtr hIMC);

        [DllImport("imm32.dll")]
        public static extern IntPtr ImmCreateContext();

        [DllImport("imm32.dll")]
        public static extern bool ImmDestroyContext(IntPtr hIMC);

        [DllImport("user32.dll")]
        public static extern IntPtr GetForegroundWindow();

        [DllImport("user32.dll")]
        public static extern IntPtr SendMessage(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam);

        [DllImport("user32.dll")]
        public static extern void keybd_event(byte bVk, byte bScan, uint dwFlags, UIntPtr dwExtraInfo);

        [DllImport("user32.dll")]
        public static extern IntPtr GetKeyboardLayout(uint idThread);

        [DllImport("user32.dll")]
        public static extern IntPtr ActivateKeyboardLayout(IntPtr hkl, uint flags);

        [DllImport("user32.dll")]
        public static extern bool PostMessage(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam);

        [DllImport("user32.dll")]
        public static extern bool RegisterHotKey(IntPtr hWnd, int id, uint fsModifiers, uint vk);

        [DllImport("user32.dll")]
        public static extern bool UnregisterHotKey(IntPtr hWnd, int id);

        [DllImport("dwmapi.dll")]
        private static extern int DwmSetWindowAttribute(IntPtr hwnd, int attr, ref int attrValue, int attrSize);

        private const int DWMWA_USE_IMMERSIVE_DARK_MODE_BEFORE_20H1 = 19;
        private const int DWMWA_USE_IMMERSIVE_DARK_MODE = 20;

        public const uint WM_HOTKEY = 0x0312;
        public const uint WM_IME_CONTROL = 0x0283;
        public const uint WM_INPUTLANGCHANGEREQUEST = 0x0050;
        public const int IMC_SETOPENSTATUS = 0x0006;
        public const byte VK_SHIFT = 0x10;
        public const uint KEYEVENTF_KEYUP = 0x0002;
        public const uint KLF_ACTIVATE = 0x00000001;
        public const int LANG_ENGLISH_US = 0x04090409;

        public delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

        public const uint GA_ROOT = 2;
        public const int VK_LBUTTON = 0x01;
        public const int VK_ESCAPE = 0x1B;
        public const int VK_F2 = 0x71;
        public const int VK_F3 = 0x72;
        public const int VK_F4 = 0x73;
        public const int VK_F6 = 0x75;
        public const uint MOUSEEVENTF_LEFTDOWN = 0x0002;
        public const uint MOUSEEVENTF_LEFTUP = 0x0004;
        
        // 热键修饰符
        public const uint MOD_NONE = 0x0000;
        public const uint MOD_ALT = 0x0001;
        public const uint MOD_CONTROL = 0x0002;
        public const uint MOD_SHIFT = 0x0004;
        public const uint MOD_WIN = 0x0008;

        [StructLayout(LayoutKind.Sequential)]
        public struct POINT
        {
            public int X;
            public int Y;
        }

        [StructLayout(LayoutKind.Sequential)]
        public struct RECT
        {
            public int Left;
            public int Top;
            public int Right;
            public int Bottom;
        }

        public static string GetWindowClassName(IntPtr hwnd)
        {
            var sb = new StringBuilder(256);
            GetClassName(hwnd, sb, sb.Capacity);
            return sb.ToString();
        }

        public static string GetWindowTitle(IntPtr hwnd)
        {
            var sb = new StringBuilder(256);
            GetWindowText(hwnd, sb, sb.Capacity);
            return sb.ToString();
        }

        public static string? GetProcessPath(IntPtr hwnd)
        {
            try
            {
                GetWindowThreadProcessId(hwnd, out uint processId);
                var process = Process.GetProcessById((int)processId);
                return process.MainModule?.FileName;
            }
            catch
            {
                return null;
            }
        }

        public static bool IsExcelWindow(IntPtr hwnd)
        {
            var className = GetWindowClassName(hwnd);
            return className.Contains("XLMAIN") || className.Contains("Excel");
        }

        /// <summary>
        /// 获取当前键盘布局
        /// </summary>
        public static IntPtr GetCurrentKeyboardLayout(IntPtr hwnd)
        {
            try
            {
                uint threadId = GetWindowThreadProcessId(hwnd, out _);
                return GetKeyboardLayout(threadId);
            }
            catch
            {
                return IntPtr.Zero;
            }
        }

        /// <summary>
        /// 切换到英文输入法（使用Python版本的方法）
        /// </summary>
        public static bool SwitchToEnglishIME(IntPtr hwnd)
        {
            try
            {
                // 方法1: 使用ActivateKeyboardLayout切换到英文
                IntPtr englishLayout = new IntPtr(LANG_ENGLISH_US);
                IntPtr result = ActivateKeyboardLayout(englishLayout, KLF_ACTIVATE);
                
                if (result != IntPtr.Zero)
                {
                    // 发送语言切换消息到目标窗口
                    PostMessage(hwnd, WM_INPUTLANGCHANGEREQUEST, IntPtr.Zero, englishLayout);
                    return true;
                }
            }
            catch { }
            
            return false;
        }

        /// <summary>
        /// 恢复到指定的键盘布局
        /// </summary>
        public static bool RestoreKeyboardLayout(IntPtr hwnd, IntPtr layout)
        {
            try
            {
                if (layout != IntPtr.Zero)
                {
                    IntPtr result = ActivateKeyboardLayout(layout, KLF_ACTIVATE);
                    if (result != IntPtr.Zero)
                    {
                        PostMessage(hwnd, WM_INPUTLANGCHANGEREQUEST, IntPtr.Zero, layout);
                        return true;
                    }
                }
            }
            catch { }
            
            return false;
        }

        /// <summary>
        /// 检查当前是否为英文输入法
        /// </summary>
        public static bool IsEnglishIME(IntPtr hwnd)
        {
            try
            {
                uint threadId = GetWindowThreadProcessId(hwnd, out _);
                IntPtr layout = GetKeyboardLayout(threadId);
                int langId = layout.ToInt32() & 0xFFFF;
                return langId == 0x0409; // 英文
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// 禁用当前窗口的输入法（使用多种方法确保成功）
        /// </summary>
        public static bool DisableIME(IntPtr hwnd)
        {
            bool wasOpen = false;
            
            try
            {
                // 首先尝试切换到英文输入法
                SwitchToEnglishIME(hwnd);

                // 方法1: 使用ImmSetOpenStatus
                IntPtr hIMC = ImmGetContext(hwnd);
                if (hIMC != IntPtr.Zero)
                {
                    wasOpen = ImmGetOpenStatus(hIMC);
                    if (wasOpen)
                    {
                        ImmSetOpenStatus(hIMC, false);
                    }
                    ImmReleaseContext(hwnd, hIMC);
                }

                // 方法2: 使用ImmAssociateContext禁用输入法上下文
                IntPtr hIMCNull = ImmAssociateContext(hwnd, IntPtr.Zero);
                if (hIMCNull != IntPtr.Zero)
                {
                    // 保存原始上下文以便恢复
                    // 这里我们不保存，因为会在EnableIME中重新关联
                }

                // 方法3: 发送WM_IME_CONTROL消息
                SendMessage(hwnd, WM_IME_CONTROL, new IntPtr(IMC_SETOPENSTATUS), IntPtr.Zero);
            }
            catch { }
            
            return wasOpen;
        }

        /// <summary>
        /// 启用当前窗口的输入法
        /// </summary>
        public static void EnableIME(IntPtr hwnd, bool restore = true, IntPtr originalLayout = default)
        {
            try
            {
                if (restore)
                {
                    // 如果有原始布局，先恢复布局
                    if (originalLayout != IntPtr.Zero)
                    {
                        RestoreKeyboardLayout(hwnd, originalLayout);
                    }

                    // 重新关联默认输入法上下文
                    IntPtr hIMC = ImmCreateContext();
                    if (hIMC != IntPtr.Zero)
                    {
                        ImmAssociateContext(hwnd, hIMC);
                        ImmSetOpenStatus(hIMC, true);
                    }
                    else
                    {
                        // 如果创建失败，尝试使用现有上下文
                        hIMC = ImmGetContext(hwnd);
                        if (hIMC != IntPtr.Zero)
                        {
                            ImmSetOpenStatus(hIMC, true);
                            ImmReleaseContext(hwnd, hIMC);
                        }
                    }
                }
            }
            catch { }
        }

        /// <summary>
        /// 获取当前输入法状态
        /// </summary>
        public static bool GetIMEStatus(IntPtr hwnd)
        {
            try
            {
                IntPtr hIMC = ImmGetContext(hwnd);
                if (hIMC != IntPtr.Zero)
                {
                    bool isOpen = ImmGetOpenStatus(hIMC);
                    ImmReleaseContext(hwnd, hIMC);
                    return isOpen;
                }
            }
            catch { }
            return false;
        }

        /// <summary>
        /// 强制切换到英文输入模式（使用多种方法）
        /// </summary>
        public static void ForceEnglishInput(IntPtr hwnd)
        {
            try
            {
                // 方法1: 使用ActivateKeyboardLayout
                SwitchToEnglishIME(hwnd);
                
                // 方法2: 模拟按下Shift键切换到英文
                keybd_event(VK_SHIFT, 0, 0, UIntPtr.Zero);
                keybd_event(VK_SHIFT, 0, KEYEVENTF_KEYUP, UIntPtr.Zero);
            }
            catch { }
        }

        /// <summary>
        /// 设置窗口标题栏为深色模式
        /// </summary>
        public static bool UseImmersiveDarkMode(IntPtr handle, bool enabled)
        {
            if (IsWindows10OrGreater(17763))
            {
                var attribute = DWMWA_USE_IMMERSIVE_DARK_MODE_BEFORE_20H1;
                if (IsWindows10OrGreater(18985))
                {
                    attribute = DWMWA_USE_IMMERSIVE_DARK_MODE;
                }

                int useImmersiveDarkMode = enabled ? 1 : 0;
                return DwmSetWindowAttribute(handle, attribute, ref useImmersiveDarkMode, sizeof(int)) == 0;
            }

            return false;
        }

        private static bool IsWindows10OrGreater(int build = -1)
        {
            var osVersion = Environment.OSVersion;
            if (osVersion.Version.Major >= 10)
            {
                if (build > 0)
                {
                    return osVersion.Version.Build >= build;
                }
                return true;
            }
            return false;
        }
    }
}
