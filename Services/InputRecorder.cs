using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using WindowInspector.Models;
using WindowInspector.Utils;

namespace WindowInspector.Services
{
    public class InputRecorder
    {
        public event EventHandler<InputRecordedEventArgs>? InputRecorded;
        public event EventHandler<string>? RecordingMessage;
        public event EventHandler? RecordingCancelled;
        public event EventHandler<List<InputPosition>>? RecordingCompleted;

        public async Task StartRecordingAsync(IntPtr targetWindow, WindowHelper.RECT windowRect, 
            int targetCount, CancellationToken cancellationToken)
        {
            var positions = new List<InputPosition>();

            await Task.Run(() =>
            {
                for (int i = 0; i < targetCount; i++)
                {
                    if (cancellationToken.IsCancellationRequested)
                    {
                        RecordingCancelled?.Invoke(this, EventArgs.Empty);
                        return;
                    }

                    RecordingMessage?.Invoke(this, $"⏳ 等待点击第 {i + 1} 个输入框...");

                    var clicked = false;
                    var startTime = DateTime.Now;
                    var lastClickState = false;

                    while ((DateTime.Now - startTime).TotalSeconds < 30)
                    {
                        if (cancellationToken.IsCancellationRequested)
                        {
                            RecordingCancelled?.Invoke(this, EventArgs.Empty);
                            return;
                        }

                        // 检测ESC键
                        if ((WindowHelper.GetAsyncKeyState(WindowHelper.VK_ESCAPE) & 0x8000) != 0)
                        {
                            RecordingCancelled?.Invoke(this, EventArgs.Empty);
                            return;
                        }

                        var currentClickState = (WindowHelper.GetAsyncKeyState(WindowHelper.VK_LBUTTON) & 0x8000) != 0;

                        if (currentClickState && !lastClickState)
                        {
                            WindowHelper.GetCursorPos(out var point);
                            var clickedHwnd = WindowHelper.WindowFromPoint(point);
                            var clickedRoot = WindowHelper.GetAncestor(clickedHwnd, WindowHelper.GA_ROOT);

                            if (clickedRoot != targetWindow)
                            {
                                RecordingMessage?.Invoke(this, $"⚠️ 点击位置不在目标窗口内，请重新点击第 {i + 1} 个输入框");
                                Thread.Sleep(500);
                                lastClickState = false;
                                continue;
                            }

                            var relX = point.X - windowRect.Left;
                            var relY = point.Y - windowRect.Top;

                            var windowWidth = windowRect.Right - windowRect.Left;
                            var windowHeight = windowRect.Bottom - windowRect.Top;

                            if (relX < 0 || relY < 0 || relX > windowWidth || relY > windowHeight)
                            {
                                RecordingMessage?.Invoke(this, $"⚠️ 点击位置超出窗口范围，请重新点击第 {i + 1} 个输入框");
                                Thread.Sleep(500);
                                lastClickState = false;
                                continue;
                            }

                            positions.Add(new InputPosition { X = relX, Y = relY });
                            InputRecorded?.Invoke(this, new InputRecordedEventArgs
                            {
                                Index = i,
                                Position = new InputPosition { X = relX, Y = relY }
                            });

                            clicked = true;
                            Thread.Sleep(300);
                            break;
                        }

                        lastClickState = currentClickState;
                        Thread.Sleep(50);
                    }

                    if (!clicked)
                    {
                        RecordingMessage?.Invoke(this, "❌ 记录超时，请重试");
                        return;
                    }
                }

                RecordingCompleted?.Invoke(this, positions);
            }, cancellationToken);
        }
    }

    public class InputRecordedEventArgs : EventArgs
    {
        public int Index { get; set; }
        public InputPosition Position { get; set; } = new();
    }
}
