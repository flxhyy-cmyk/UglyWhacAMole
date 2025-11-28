using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using WindowInspector.Models;
using WindowInspector.Utils;
using Emgu.CV;

namespace WindowInspector.Services
{

    
    /// <summary>
    /// 打地鼠服务 - 使用 Python OpenCV 进行图像识别和自动点击
    /// </summary>
    public class MoleHunter : IDisposable
    {
        private bool _isRunning = false;
        private bool _continuousClick = false;
        private bool _fullScreenMatch = false;
        private CancellationTokenSource? _cts;
        private EmguImageMatcher _emguMatcher;
        private bool _disposed = false;
        private List<MoleGroup>? _allMoleGroups; // 保存所有分组以支持跳转
        
        public event EventHandler<string>? LogMessage;
        public event EventHandler<MoleFoundEventArgs>? MoleFound;
        public event EventHandler? HuntingStopped;
        
        public MoleHunter()
        {
            _emguMatcher = new EmguImageMatcher();
        }

        /// <summary>
        /// 设置是否持续点击直到目标消失
        /// </summary>
        public void SetContinuousClick(bool enabled)
        {
            _continuousClick = enabled;
            LogMessage?.Invoke(this, $"⚙️ 持续点击模式: {(enabled ? "已启用" : "已禁用")}");
        }
        
        /// <summary>
        /// 设置是否全图匹配模式
        /// </summary>
        public void SetFullScreenMatch(bool enabled)
        {
            _fullScreenMatch = enabled;
            LogMessage?.Invoke(this, $"⚙️ 全图匹配模式: {(enabled ? "已启用" : "已禁用")}");
        }
        
        /// <summary>
        /// 开始打地鼠
        /// </summary>
        public void Start(List<MoleItem> moles, List<Point>? idleClickPositions = null, List<MoleGroup>? allMoleGroups = null)
        {
            if (_isRunning)
                return;
            
            _isRunning = true;
            _cts = new CancellationTokenSource();
            _allMoleGroups = allMoleGroups; // 保存所有分组
            
            Task.Run(() => HuntingLoop(moles, idleClickPositions, _cts.Token));
            LogMessage?.Invoke(this, "🎯 打地鼠已启动 (使用 Emgu.CV 原生识图)");
        }
        
        /// <summary>
        /// 停止打地鼠
        /// </summary>
        public void Stop()
        {
            if (!_isRunning)
                return;
                
            _isRunning = false;
            _cts?.Cancel();
            LogMessage?.Invoke(this, "⏸️ 打地鼠已停止");
            HuntingStopped?.Invoke(this, EventArgs.Empty);
        }
        
        /// <summary>
        /// 清空图像模板缓存
        /// </summary>
        public void ClearTemplateCache()
        {
            _emguMatcher?.ClearTemplateCache();
        }
        
        public bool IsRunning => _isRunning;
        
        private async Task HuntingLoop(List<MoleItem> moles, List<Point>? idleClickPositions, CancellationToken token)
        {
            try
            {
                while (!token.IsCancellationRequested)
                {
                    if (_fullScreenMatch)
                    {
                        // 全图匹配模式
                        await FullScreenMatchLoop(moles, idleClickPositions, token);
                    }
                    else
                    {
                        // 顺序匹配模式（原逻辑）
                        await SequentialMatchLoop(moles, idleClickPositions, token);
                    }
                    
                    // 一轮结束，短暂延迟后开始下一轮
                    await Task.Delay(100, token);
                }
            }
            catch (OperationCanceledException)
            {
                // 正常取消
            }
            catch (Exception ex)
            {
                LogMessage?.Invoke(this, $"❌ 错误: {ex.Message}");
            }
            finally
            {
                _isRunning = false;
            }
        }
        
        /// <summary>
        /// 全图匹配模式：一次性识别所有截图，找到就点击，没找到就执行空击
        /// </summary>
        private async Task FullScreenMatchLoop(List<MoleItem> moles, List<Point>? idleClickPositions, CancellationToken token)
        {
            // 获取所有启用的截图地鼠（排除空击地鼠）
            var screenshotMoles = moles.Where(m => m.IsEnabled && !m.IsIdleClick && !string.IsNullOrEmpty(m.ImagePath)).ToList();
            
            if (screenshotMoles.Count == 0)
            {
                LogMessage?.Invoke(this, "⚠️ 没有启用的截图地鼠");
                await Task.Delay(1000, token);
                return;
            }
            
            LogMessage?.Invoke(this, $"📸 全图识别中...");
            
            // 截取屏幕并转换为 Mat（只转换一次，提高性能）
            using var screenshot = CaptureScreen();
            using var screenshotMat = _emguMatcher.CreateMatFromBitmap(screenshot);
            
            // 并行识别所有地鼠
            var matchTasks = screenshotMoles.Select(async mole =>
            {
                var result = await Task.Run(() => _emguMatcher.FindTemplate(screenshotMat, mole.ImagePath, mole.SimilarityThreshold));
                return new { Mole = mole, Result = result };
            }).ToList();
            
            var matchResults = await Task.WhenAll(matchTasks);
            
            // 筛选出成功匹配的结果
            var foundMatches = matchResults
                .Where(r => r.Result != null && r.Result.Found)
                .Select(r => new
                {
                    r.Mole,
                    r.Result,
                    r.Result.Confidence
                })
                .ToList();
            
            if (foundMatches.Count > 0)
            {
                // 找到匹配项，按列表顺序点击
                LogMessage?.Invoke(this, $"✅ 全图识别完成，找到 {foundMatches.Count} 个匹配项");
                
                // 按原始列表顺序排序
                var orderedMatches = foundMatches
                    .OrderBy(m => screenshotMoles.IndexOf(m.Mole))
                    .ToList();
                
                foreach (var match in orderedMatches)
                {
                    if (token.IsCancellationRequested) break;
                    
                    // 点击匹配项
                    ClickAt(match.Result.Center);
                    MoleFound?.Invoke(this, new MoleFoundEventArgs(match.Mole.Name, match.Result.Center));
                    LogMessage?.Invoke(this, $"🎯 [{match.Mole.Name}] 点击 ({match.Result.Center.X},{match.Result.Center.Y}) | 置信度:{match.Confidence:F2} (阈值:{match.Mole.SimilarityThreshold:F2})");
                    
                    // 点击间隔
                    await Task.Delay(100, token);
                }
            }
            else
            {
                // 没有找到任何匹配，执行空击步骤
                LogMessage?.Invoke(this, "⏭️ 全图识别无匹配，执行空击步骤");
                
                if (idleClickPositions != null && idleClickPositions.Count > 0)
                {
                    foreach (var pos in idleClickPositions)
                    {
                        if (token.IsCancellationRequested) break;
                        
                        ClickAt(pos);
                        LogMessage?.Invoke(this, $"💤 空击 ({pos.X},{pos.Y})");
                        await Task.Delay(50, token);
                    }
                }
                else
                {
                    LogMessage?.Invoke(this, "⚠️ 未设置空击位置");
                    await Task.Delay(500, token);
                }
            }
        }
        
        /// <summary>
        /// 顺序匹配模式：按列表顺序逐个检查地鼠（原逻辑）
        /// </summary>
        private async Task SequentialMatchLoop(List<MoleItem> moles, List<Point>? idleClickPositions, CancellationToken token)
        {
            await ExecuteMoleSequence(moles, idleClickPositions, token);
        }

        /// <summary>
        /// 执行地鼠序列（支持跳转）
        /// </summary>
        private async Task ExecuteMoleSequence(List<MoleItem> moles, List<Point>? idleClickPositions, CancellationToken token)
        {
            await ExecuteMoleSequenceInternal(moles, idleClickPositions, token, moles.Count, 0);
        }

        /// <summary>
        /// 内部递归执行地鼠序列（支持跳转）
        /// </summary>
        private async Task ExecuteMoleSequenceInternal(List<MoleItem> moles, List<Point>? idleClickPositions, CancellationToken token, int totalSteps, int startIndex = 0)
        {
            int currentStep = 0;
            
            // 按列表顺序逐个检查地鼠
            for (int i = startIndex; i < moles.Count; i++)
            {
                var mole = moles[i];
                currentStep++;
                
                if (!mole.IsEnabled || token.IsCancellationRequested)
                    continue;
                
                // 如果是跳转步骤
                if (mole.IsJump)
                {
                    LogMessage?.Invoke(this, $"[{currentStep}/{totalSteps}] 🔗 跳转到 {mole.JumpTargetGroup}");
                    
                    // 查找目标分组
                    if (_allMoleGroups != null)
                    {
                        var targetGroup = _allMoleGroups.FirstOrDefault(g => g.Name == mole.JumpTargetGroup);
                        if (targetGroup != null)
                        {
                            // 确定起始步骤
                            int targetStartIndex = mole.JumpTargetStep >= 0 ? mole.JumpTargetStep : 0;
                            
                            if (targetStartIndex < targetGroup.Moles.Count)
                            {
                                // 执行目标分组的步骤
                                if (mole.JumpTargetStep >= 0)
                                {
                                    LogMessage?.Invoke(this, $"📂 进入分组: {targetGroup.Name} (从步骤 {targetStartIndex + 1} 开始)");
                                }
                                else
                                {
                                    LogMessage?.Invoke(this, $"📂 进入分组: {targetGroup.Name}");
                                }
                                
                                await ExecuteMoleSequenceInternal(targetGroup.Moles, targetGroup.IdleClickPositions, token, totalSteps, targetStartIndex);
                                LogMessage?.Invoke(this, $"📂 返回分组");
                            }
                            else
                            {
                                LogMessage?.Invoke(this, $"⚠️ 目标步骤索引超出范围: {targetStartIndex}");
                            }
                        }
                        else
                        {
                            LogMessage?.Invoke(this, $"⚠️ 未找到目标分组: {mole.JumpTargetGroup}");
                        }
                    }
                    
                    await Task.Delay(50, token);
                    continue;
                }
                
                // 如果是空击地鼠
                if (mole.IsIdleClick && mole.IdleClickPosition.HasValue)
                {
                    // 检查是否设置了停止打地鼠
                    if (mole.StopHunting)
                    {
                        LogMessage?.Invoke(this, $"[{currentStep}/{totalSteps}] ⏹️ 执行到停止步骤，打地鼠已停止");
                        Stop(); // 停止打地鼠
                        return; // 退出执行
                    }
                    
                    // 执行一次空击
                    ClickAt(mole.IdleClickPosition.Value);
                    LogMessage?.Invoke(this, $"[{currentStep}/{totalSteps}] 空击地鼠打击 ({mole.IdleClickPosition.Value.X}, {mole.IdleClickPosition.Value.Y})");
                    // 跳到下一个地鼠
                    await Task.Delay(50, token);
                    continue;
                }
                
                // 如果是截图地鼠
                if (!mole.IsIdleClick && !string.IsNullOrEmpty(mole.ImagePath))
                {
                    // 如果启用了"持续等待直到出现"
                    if (mole.WaitUntilAppear)
                    {
                        LogMessage?.Invoke(this, $"[{currentStep}/{totalSteps}] ⏳ 等待图像出现: {mole.Name}");
                        
                        ImageMatchResult? matchResult = null;
                        int waitCount = 0;
                        
                        // 持续扫描直到找到图像
                        while (!token.IsCancellationRequested)
                        {
                            matchResult = FindImageWithEmgu(mole.ImagePath, mole.SimilarityThreshold);
                            
                            if (matchResult != null && matchResult.Found)
                            {
                                // 找到了，退出等待循环
                                LogMessage?.Invoke(this, $"[{currentStep}/{totalSteps}] ✅ 图像已出现，匹配阈值:{matchResult.Confidence:F2} (等待了 {waitCount} 次扫描)");
                                break;
                            }
                            
                            waitCount++;
                            if (waitCount % 10 == 0)
                            {
                                LogMessage?.Invoke(this, $"[{currentStep}/{totalSteps}] ⏳ 继续等待... (已扫描 {waitCount} 次)");
                            }
                            
                            // 等待一小段时间后再次扫描
                            await Task.Delay(100, token);
                        }
                        
                        // 找到后点击
                        if (matchResult != null && matchResult.Found)
                        {
                            ClickAt(matchResult.Center);
                            MoleFound?.Invoke(this, new MoleFoundEventArgs(mole.Name, matchResult.Center));
                            LogMessage?.Invoke(this, $"[{currentStep}/{totalSteps}] 🎯 截图地鼠打击成功 ({matchResult.Center.X}, {matchResult.Center.Y}) | 置信度:{matchResult.Confidence:F2} (阈值:{mole.SimilarityThreshold:F2})");
                            
                            // 如果启用了"持续点击直到消失"
                            if (mole.ClickUntilDisappear)
                            {
                                int clickCount = 1;
                                while (!token.IsCancellationRequested)
                                {
                                    // 等待 200ms
                                    await Task.Delay(200, token);
                                    
                                    // 再次检查目标是否还存在
                                    var recheckResult = FindImageWithEmgu(mole.ImagePath, mole.SimilarityThreshold);
                                    
                                    if (recheckResult != null && recheckResult.Found)
                                    {
                                        // 目标仍然存在，继续点击
                                        clickCount++;
                                        ClickAt(recheckResult.Center);
                                        LogMessage?.Invoke(this, $"[{currentStep}/{totalSteps}] 🔄 持续点击第 {clickCount} 次 ({recheckResult.Center.X}, {recheckResult.Center.Y}) | 置信度:{recheckResult.Confidence:F2}");
                                    }
                                    else
                                    {
                                        // 目标已消失，退出循环
                                        LogMessage?.Invoke(this, $"[{currentStep}/{totalSteps}] ✅ 图像已消失，共点击 {clickCount} 次");
                                        break;
                                    }
                                }
                            }
                            
                            // 如果启用了"点击后等待"
                            if (mole.WaitAfterClick && mole.WaitAfterClickMs > 0)
                            {
                                LogMessage?.Invoke(this, $"[{currentStep}/{totalSteps}] ⏱️ 等待 {mole.WaitAfterClickMs} ms...");
                                await Task.Delay(mole.WaitAfterClickMs, token);
                            }
                        }
                    }
                    else
                    {
                        // 正常模式：扫描一次
                        var matchResult = FindImageWithEmgu(mole.ImagePath, mole.SimilarityThreshold);
                        
                        if (matchResult != null && matchResult.Found)
                        {
                            // 找到地鼠，点击中心点
                            ClickAt(matchResult.Center);
                            MoleFound?.Invoke(this, new MoleFoundEventArgs(mole.Name, matchResult.Center));
                            LogMessage?.Invoke(this, $"[{currentStep}/{totalSteps}] 🎯 截图地鼠打击成功 ({matchResult.Center.X}, {matchResult.Center.Y}) | 置信度:{matchResult.Confidence:F2} (阈值:{mole.SimilarityThreshold:F2})");
                            
                            // 如果启用了"持续点击直到消失"（针对当前地鼠）
                            if (mole.ClickUntilDisappear)
                            {
                                int clickCount = 1;
                                while (!token.IsCancellationRequested)
                                {
                                    // 等待 200ms
                                    await Task.Delay(200, token);
                                    
                                    // 再次检查目标是否还存在
                                    var recheckResult = FindImageWithEmgu(mole.ImagePath, mole.SimilarityThreshold);
                                    
                                    if (recheckResult != null && recheckResult.Found)
                                    {
                                        // 目标仍然存在，继续点击
                                        clickCount++;
                                        ClickAt(recheckResult.Center);
                                        LogMessage?.Invoke(this, $"[{currentStep}/{totalSteps}] 🔄 持续点击第 {clickCount} 次 ({recheckResult.Center.X}, {recheckResult.Center.Y}) | 置信度:{recheckResult.Confidence:F2}");
                                    }
                                    else
                                    {
                                        // 目标已消失，退出循环
                                        LogMessage?.Invoke(this, $"[{currentStep}/{totalSteps}] ✅ 图像已消失，共点击 {clickCount} 次");
                                        break;
                                    }
                                }
                            }
                            // 如果启用了全局持续点击模式（旧功能，保持兼容）
                            else if (_continuousClick)
                            {
                                int clickCount = 1;
                                while (!token.IsCancellationRequested)
                                {
                                    // 等待 200ms
                                    await Task.Delay(200, token);
                                    
                                    // 再次检查目标是否还存在
                                    var recheckResult = FindImageWithEmgu(mole.ImagePath, mole.SimilarityThreshold);
                                    
                                    if (recheckResult != null && recheckResult.Found)
                                    {
                                        // 目标仍然存在，继续点击
                                        clickCount++;
                                        ClickAt(recheckResult.Center);
                                    }
                                    else
                                    {
                                        // 目标已消失，退出循环
                                        break;
                                    }
                                }
                            }
                            
                            // 如果启用了"点击后等待"
                            if (mole.WaitAfterClick && mole.WaitAfterClickMs > 0)
                            {
                                LogMessage?.Invoke(this, $"[{currentStep}/{totalSteps}] ⏱️ 等待 {mole.WaitAfterClickMs} ms...");
                                await Task.Delay(mole.WaitAfterClickMs, token);
                            }
                        }
                        else
                        {
                            // 未找到地鼠
                            if (mole.JumpToPreviousOnFail && i > startIndex)
                            {
                                // 启用了"识别失败跳转到上一步"，且不是第一步
                                LogMessage?.Invoke(this, $"[{currentStep}/{totalSteps}] ⚠️ 截图地鼠未找到，跳转到上一个步骤");
                                i = i - 2; // -2 是因为循环会 +1，所以实际是回到上一步
                                currentStep--; // 步骤计数也要回退
                                await Task.Delay(50, token);
                                continue;
                            }
                            else
                            {
                                // 未找到地鼠，跳过此步骤
                                string confidenceInfo = matchResult != null ? $" | 最高置信度:{matchResult.Confidence:F2} (阈值:{mole.SimilarityThreshold:F2})" : "";
                                LogMessage?.Invoke(this, $"[{currentStep}/{totalSteps}] ⏭️ 截图地鼠未找到 (跳过){confidenceInfo}");
                            }
                        }
                    }
                    
                    // 短暂延迟后继续下一个步骤
                    await Task.Delay(50, token);
                }
            }
        }
        
        /// <summary>
        /// 捕获整个屏幕
        /// </summary>
        private Bitmap CaptureScreen()
        {
            var bounds = Screen.PrimaryScreen.Bounds;
            var bitmap = new Bitmap(bounds.Width, bounds.Height);
            
            using (var g = Graphics.FromImage(bitmap))
            {
                g.CopyFromScreen(Point.Empty, Point.Empty, bounds.Size);
            }
            
            return bitmap;
        }
        
        /// <summary>
        /// 使用 Emgu.CV 进行图像匹配
        /// </summary>
        private ImageMatchResult? FindImageWithEmgu(string templatePath, double threshold)
        {
            try
            {
                using var screenshot = CaptureScreen();
                var result = _emguMatcher.FindTemplate(screenshot, templatePath, threshold);
                
                if (result != null && !string.IsNullOrEmpty(result.Error))
                {
                    LogMessage?.Invoke(this, $"❌ 识图错误: {result.Error}");
                }
                
                return result;
            }
            catch (Exception ex)
            {
                LogMessage?.Invoke(this, $"❌ 识图异常: {ex.Message}");
                return null;
            }
        }
        
        /// <summary>
        /// 在指定位置点击鼠标
        /// </summary>
        private void ClickAt(Point location)
        {
            // 保存当前鼠标位置
            WindowHelper.GetCursorPos(out var oldPos);
            
            // 移动到目标位置
            WindowHelper.SetCursorPos(location.X, location.Y);
            
            // 模拟鼠标点击
            WindowHelper.mouse_event(WindowHelper.MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0);
            Thread.Sleep(10);
            WindowHelper.mouse_event(WindowHelper.MOUSEEVENTF_LEFTUP, 0, 0, 0, 0);
            
            // 恢复鼠标位置（可选）
            // WindowHelper.SetCursorPos(oldPos.X, oldPos.Y);
        }

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        protected virtual void Dispose(bool disposing)
        {
            if (!_disposed)
            {
                if (disposing)
                {
                    Stop();
                    _emguMatcher?.Dispose();
                    _cts?.Dispose();
                }
                _disposed = true;
            }
        }
    }
    
    public class MoleFoundEventArgs : EventArgs
    {
        public string MoleName { get; }
        public Point Location { get; }
        
        public MoleFoundEventArgs(string moleName, Point location)
        {
            MoleName = moleName;
            Location = location;
        }
    }
}
