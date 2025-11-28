using System;
using System.Threading;
using System.Threading.Tasks;
using WindowInspector.Utils;

namespace WindowInspector.Services
{
    public class WindowSelector
    {
        public event EventHandler<WindowSelectedEventArgs>? WindowSelected;
        public event EventHandler<string>? SelectionTimeout;

        public async Task StartSelectionAsync(CancellationToken cancellationToken)
        {
            await Task.Run(() =>
            {
                var startTime = DateTime.Now;
                var lastClickState = false;

                while ((DateTime.Now - startTime).TotalSeconds < 10)
                {
                    if (cancellationToken.IsCancellationRequested)
                        return;

                    var currentClickState = (WindowHelper.GetAsyncKeyState(WindowHelper.VK_LBUTTON) & 0x8000) != 0;

                    if (currentClickState && !lastClickState)
                    {
                        WindowHelper.GetCursorPos(out var point);
                        var hwnd = WindowHelper.WindowFromPoint(point);
                        var rootHwnd = WindowHelper.GetAncestor(hwnd, WindowHelper.GA_ROOT);

                        if (rootHwnd != IntPtr.Zero)
                        {
                            WindowHelper.GetWindowRect(rootHwnd, out var rect);
                            var windowTitle = WindowHelper.GetWindowTitle(rootHwnd);
                            var windowClass = WindowHelper.GetWindowClassName(rootHwnd);

                            WindowSelected?.Invoke(this, new WindowSelectedEventArgs
                            {
                                WindowHandle = rootHwnd,
                                WindowRect = rect,
                                WindowTitle = windowTitle,
                                WindowClass = windowClass
                            });
                            return;
                        }
                    }

                    lastClickState = currentClickState;
                    Thread.Sleep(50);
                }

                SelectionTimeout?.Invoke(this, "窗口选择超时，请重试");
            }, cancellationToken);
        }
    }

    public class WindowSelectedEventArgs : EventArgs
    {
        public IntPtr WindowHandle { get; set; }
        public WindowHelper.RECT WindowRect { get; set; }
        public string WindowTitle { get; set; } = string.Empty;
        public string WindowClass { get; set; } = string.Empty;
    }
}
