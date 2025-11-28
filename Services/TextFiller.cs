using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using WindowInspector.Models;
using WindowInspector.Utils;

namespace WindowInspector.Services
{
    public class TextFiller
    {
        public async Task FillTextAsync(IntPtr targetWindow, WindowHelper.RECT windowRect,
            List<InputPosition> positions, List<string> texts)
        {
            await Task.Run(() =>
            {
                // 激活目标窗口
                WindowHelper.SetForegroundWindow(targetWindow);
                Thread.Sleep(200);

                // 记住原始键盘布局
                IntPtr originalLayout = WindowHelper.GetCurrentKeyboardLayout(targetWindow);

                // 禁用输入法并确认
                var (success, imeWasEnabled) = DisableIMEAndConfirm(targetWindow);
                
                if (!success)
                {
                    throw new Exception("无法禁用输入法，请手动切换到英文输入模式后重试");
                }

                try
                {
                    for (int i = 0; i < Math.Min(positions.Count, texts.Count); i++)
                    {
                        var pos = positions[i];
                        var text = texts[i];

                        // 先确认输入法已禁用（在移动鼠标之前）
                        if (!EnsureIMEDisabled(targetWindow))
                        {
                            throw new Exception($"填充第{i + 1}个输入框前输入法无法禁用，已停止填充");
                        }

                        // 等待输入法切换完全稳定，避免误输入
                        Thread.Sleep(300);

                        // 输入法切换完成后，再移动鼠标
                        WindowHelper.GetWindowRect(targetWindow, out var currentRect);
                        var absX = currentRect.Left + pos.X;
                        var absY = currentRect.Top + pos.Y;

                        WindowHelper.GetCursorPos(out var originalPos);
                        WindowHelper.SetCursorPos(absX, absY);
                        Thread.Sleep(150);

                        // 双击输入框
                        // 第一次点击
                        WindowHelper.mouse_event(WindowHelper.MOUSEEVENTF_LEFTDOWN, 0, 0, 0, UIntPtr.Zero);
                        WindowHelper.mouse_event(WindowHelper.MOUSEEVENTF_LEFTUP, 0, 0, 0, UIntPtr.Zero);
                        Thread.Sleep(100);
                        
                        // 第二次点击
                        WindowHelper.mouse_event(WindowHelper.MOUSEEVENTF_LEFTDOWN, 0, 0, 0, UIntPtr.Zero);
                        WindowHelper.mouse_event(WindowHelper.MOUSEEVENTF_LEFTUP, 0, 0, 0, UIntPtr.Zero);
                        Thread.Sleep(200);

                        // 双击后再次确认输入法状态，避免双击过程中输入法被激活
                        if (WindowHelper.GetIMEStatus(targetWindow))
                        {
                            WindowHelper.DisableIME(targetWindow);
                            Thread.Sleep(100);
                        }

                        // 清空可能的键盘缓冲区，等待所有按键事件处理完毕
                        System.Windows.Forms.Application.DoEvents();
                        Thread.Sleep(150);

                        // 直接输入文本（不使用Ctrl+A全选，避免产生额外按键）
                        System.Windows.Forms.SendKeys.SendWait(text);
                        Thread.Sleep(200);

                        WindowHelper.SetCursorPos(originalPos.X, originalPos.Y);
                        Thread.Sleep(100); // 每个输入框之间增加延迟
                    }
                }
                finally
                {
                    // 恢复原始键盘布局和输入法状态
                    if (originalLayout != IntPtr.Zero)
                    {
                        WindowHelper.RestoreKeyboardLayout(targetWindow, originalLayout);
                        Thread.Sleep(100);
                    }

                    if (imeWasEnabled)
                    {
                        WindowHelper.EnableIME(targetWindow, true, originalLayout);
                    }
                }
            });
        }

        /// <summary>
        /// 禁用输入法并确认已禁用
        /// </summary>
        /// <returns>(是否成功禁用, 原始是否启用)</returns>
        private (bool success, bool wasEnabled) DisableIMEAndConfirm(IntPtr hwnd)
        {
            // 记录原始状态
            bool wasEnabled = WindowHelper.GetIMEStatus(hwnd);

            // 尝试禁用输入法，最多重试10次
            for (int i = 0; i < 10; i++)
            {
                // 先切换到英文输入法
                WindowHelper.SwitchToEnglishIME(hwnd);
                Thread.Sleep(100);

                // 禁用IME
                WindowHelper.DisableIME(hwnd);
                Thread.Sleep(100);

                // 不使用ForceEnglishInput，避免Shift键产生误输入
                // WindowHelper.ForceEnglishInput(hwnd);
                // Thread.Sleep(50);

                // 确认是否已切换到英文且禁用
                if (WindowHelper.IsEnglishIME(hwnd) && !WindowHelper.GetIMEStatus(hwnd))
                {
                    // 再等待一下确保稳定
                    Thread.Sleep(100);
                    
                    // 最后再确认一次
                    if (WindowHelper.IsEnglishIME(hwnd) && !WindowHelper.GetIMEStatus(hwnd))
                    {
                        return (true, wasEnabled); // 成功禁用
                    }
                }

                Thread.Sleep(100);
            }

            // 如果10次后仍未禁用，最后尝试一次（不使用键盘模拟）
            WindowHelper.SwitchToEnglishIME(hwnd);
            WindowHelper.DisableIME(hwnd);
            // 不使用ForceEnglishInput，避免Shift键产生误输入
            // WindowHelper.ForceEnglishInput(hwnd);
            Thread.Sleep(200);
            
            // 最终确认
            bool finallyDisabled = !WindowHelper.GetIMEStatus(hwnd);
            return (finallyDisabled, wasEnabled);
        }

        /// <summary>
        /// 确保输入法已禁用（不产生额外输入）
        /// </summary>
        /// <returns>是否成功保持禁用状态</returns>
        private bool EnsureIMEDisabled(IntPtr hwnd)
        {
            // 多次检查和禁用，但不使用可能产生输入的方法
            for (int i = 0; i < 5; i++)
            {
                // 检查是否为英文输入法且已禁用
                if (!WindowHelper.IsEnglishIME(hwnd) || WindowHelper.GetIMEStatus(hwnd))
                {
                    // 只使用API方法，不使用键盘模拟（避免ForceEnglishInput）
                    WindowHelper.SwitchToEnglishIME(hwnd);
                    WindowHelper.DisableIME(hwnd);
                    Thread.Sleep(150);
                }
                else
                {
                    return true; // 已经是英文且禁用，成功
                }
            }
            
            // 5次后仍然无法禁用
            return WindowHelper.IsEnglishIME(hwnd) && !WindowHelper.GetIMEStatus(hwnd);
        }

        [DllImport("oleaut32.dll", PreserveSig = false)]
        private static extern void GetActiveObject(ref Guid rclsid, IntPtr pvReserved, [MarshalAs(UnmanagedType.IUnknown)] out object ppunk);

        public async Task FillExcelCellsAsync(List<string> cells, List<string> texts)
        {
            await Task.Run(() =>
            {
                try
                {
                    var excelType = Type.GetTypeFromProgID("Excel.Application");
                    if (excelType == null)
                        throw new Exception("未检测到Excel应用程序");

                    // 尝试获取运行中的Excel实例
                    object? excelApp = null;
                    try
                    {
                        var clsid = new Guid("00024500-0000-0000-C000-000000000046"); // Excel.Application CLSID
                        GetActiveObject(ref clsid, IntPtr.Zero, out excelApp);
                    }
                    catch
                    {
                        throw new Exception("没有运行中的Excel实例，请先打开Excel");
                    }

                    if (excelApp == null)
                        throw new Exception("无法连接到Excel");

                    dynamic excel = excelApp;
                    dynamic workbook = excel.ActiveWorkbook;
                    if (workbook == null)
                        throw new Exception("没有打开的Excel工作簿");

                    dynamic worksheet = workbook.ActiveSheet;
                    if (worksheet == null)
                        throw new Exception("没有活动的工作表");

                    for (int i = 0; i < Math.Min(cells.Count, texts.Count); i++)
                    {
                        dynamic range = worksheet.Range[cells[i]];
                        range.Value = texts[i];
                    }
                }
                catch (Exception ex)
                {
                    throw new Exception($"填充Excel失败: {ex.Message}");
                }
            });
        }
    }
}
