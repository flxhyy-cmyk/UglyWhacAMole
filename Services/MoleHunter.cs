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
        public event EventHandler<string>? OnConfigSwitchRequested;
        public event EventHandler<string>? OnTextContentSwitchRequested;
        
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
        
        private string? _currentGroupName; // 当前分组名称
        
        /// <summary>
        /// 开始打地鼠
        /// </summary>
        public void Start(List<MoleItem> moles, List<MoleGroup>? allMoleGroups = null, string? groupName = null)
        {
            if (_isRunning)
                return;
            
            _isRunning = true;
            _cts = new CancellationTokenSource();
            _allMoleGroups = allMoleGroups; // 保存所有分组
            _currentGroupName = groupName; // 保存当前分组名称
            
            Task.Run(() => HuntingLoop(moles, _cts.Token));
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
        
        private async Task HuntingLoop(List<MoleItem> moles, CancellationToken token)
        {
            try
            {
                while (!token.IsCancellationRequested)
                {
                    if (_fullScreenMatch)
                    {
                        // 全图匹配模式
                        await FullScreenMatchLoop(moles, token);
                    }
                    else
                    {
                        // 顺序匹配模式（原逻辑）
                        await SequentialMatchLoop(moles, token);
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
        private async Task FullScreenMatchLoop(List<MoleItem> moles, CancellationToken token)
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
                    LogMessage?.Invoke(this, $"🎯 [{match.Mole.Name}] 点击 ({match.Result.Center.X},{match.Result.Center.Y})");
                    
                    // 点击间隔
                    await Task.Delay(100, token);
                }
            }
            else
            {
                // 没有找到任何匹配，执行空击步骤
                LogMessage?.Invoke(this, "⏭️ 全图识别无匹配，执行空击步骤");
                
                // 从moles列表中获取所有启用的空击步骤
                var idleClickMoles = moles.Where(m => m.IsEnabled && m.IsIdleClick && m.IdleClickPosition.HasValue).ToList();
                
                if (idleClickMoles.Count > 0)
                {
                    foreach (var idleMole in idleClickMoles)
                    {
                        if (token.IsCancellationRequested) break;
                        
                        ClickAt(idleMole.IdleClickPosition.Value);
                        LogMessage?.Invoke(this, $"💤 空击 ({idleMole.IdleClickPosition.Value.X},{idleMole.IdleClickPosition.Value.Y})");
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
        private async Task SequentialMatchLoop(List<MoleItem> moles, CancellationToken token)
        {
            await ExecuteMoleSequence(moles, token);
        }

        /// <summary>
        /// 执行地鼠序列（支持跳转）
        /// </summary>
        private async Task ExecuteMoleSequence(List<MoleItem> moles, CancellationToken token)
        {
            await ExecuteMoleSequenceInternal(moles, token, moles.Count, 0);
        }

        /// <summary>
        /// 内部递归执行地鼠序列（支持跳转）
        /// </summary>
        private async Task ExecuteMoleSequenceInternal(List<MoleItem> moles, CancellationToken token, int totalSteps, int startIndex = 0)
        {
            int currentStep = 0;
            string groupPrefix = string.IsNullOrEmpty(_currentGroupName) ? "" : $"[{_currentGroupName}]";
            
            // 按列表顺序逐个检查地鼠
            for (int i = startIndex; i < moles.Count; i++)
            {
                var mole = moles[i];
                currentStep++;
                
                if (!mole.IsEnabled || token.IsCancellationRequested)
                    continue;
                
                string stepPrefix = $"{groupPrefix}[{currentStep}/{totalSteps}]";
                
                // 如果是配置步骤
                if (mole.IsConfigStep)
                {
                    LogMessage?.Invoke(this, $"{stepPrefix} ⚙️ 配置步骤: {mole.Name}");
                    
                    // 执行配置切换
                    if (mole.SwitchConfig)
                    {
                        try
                        {
                            OnConfigSwitchRequested?.Invoke(this, mole.TargetConfigName);
                            LogMessage?.Invoke(this, $"{stepPrefix} ✅ 已切换配置: {mole.TargetConfigName}");
                            
                            if (mole.ConfigSwitchWaitMs > 0)
                            {
                                await Task.Delay(mole.ConfigSwitchWaitMs, token);
                                LogMessage?.Invoke(this, $"{stepPrefix} ⏱️ 已等待 {mole.ConfigSwitchWaitMs}ms");
                            }
                        }
                        catch (Exception ex)
                        {
                            LogMessage?.Invoke(this, $"{stepPrefix} ❌ 配置切换失败: {ex.Message}");
                        }
                    }
                    
                    // 执行填充内容切换
                    if (mole.SwitchTextContent)
                    {
                        try
                        {
                            OnTextContentSwitchRequested?.Invoke(this, mole.TargetTextName);
                            LogMessage?.Invoke(this, $"{stepPrefix} ✅ 已切换填充内容: {mole.TargetTextName}");
                            
                            if (mole.TextSwitchWaitMs > 0)
                            {
                                await Task.Delay(mole.TextSwitchWaitMs, token);
                                LogMessage?.Invoke(this, $"{stepPrefix} ⏱️ 已等待 {mole.TextSwitchWaitMs}ms");
                            }
                        }
                        catch (Exception ex)
                        {
                            LogMessage?.Invoke(this, $"{stepPrefix} ❌ 填充内容切换失败: {ex.Message}");
                        }
                    }
                    
                    await Task.Delay(50, token);
                    continue;
                }
                
                // 如果是跳转步骤
                if (mole.IsJump)
                {
                    // 检查是否为键盘按键输入模式
                    if (mole.SendKeyPress)
                    {
                        // 键盘按键输入模式
                        bool hasKeyPress = !string.IsNullOrEmpty(mole.KeyPressDefinition);
                        bool hasMouseScroll = mole.EnableMouseScroll;
                        
                        // 如果键盘按键先执行
                        if (hasKeyPress && mole.IsKeyPressExecuteFirst)
                        {
                            LogMessage?.Invoke(this, $"{stepPrefix} ⌨️ 发送按键: {mole.KeyPressDefinition}");
                            
                            try
                            {
                                SendKeyPress(mole.KeyPressDefinition);
                                LogMessage?.Invoke(this, $"{stepPrefix} ✅ 按键已发送");
                                
                                // 等待指定时间
                                if (mole.KeyPressWaitMs > 0)
                                {
                                    await Task.Delay(mole.KeyPressWaitMs, token);
                                    LogMessage?.Invoke(this, $"{stepPrefix} ⏱️ 已等待 {mole.KeyPressWaitMs}ms");
                                }
                            }
                            catch (Exception ex)
                            {
                                LogMessage?.Invoke(this, $"{stepPrefix} ❌ 按键发送失败: {ex.Message}");
                            }
                            
                            // 然后执行鼠标滚动
                            if (hasMouseScroll)
                            {
                                var direction = mole.ScrollUp ? "向上" : "向下";
                                LogMessage?.Invoke(this, $"{stepPrefix} 🖱️ 鼠标{direction}滚动 {mole.ScrollCount} 次");
                                
                                try
                                {
                                    PerformMouseScroll(mole.ScrollUp, mole.ScrollCount);
                                    LogMessage?.Invoke(this, $"{stepPrefix} ✅ 滚动完成");
                                    
                                    // 滚动后等待
                                    if (mole.ScrollWaitMs > 0)
                                    {
                                        await Task.Delay(mole.ScrollWaitMs, token);
                                        LogMessage?.Invoke(this, $"{stepPrefix} ⏱️ 滚动后已等待 {mole.ScrollWaitMs}ms");
                                    }
                                }
                                catch (Exception ex)
                                {
                                    LogMessage?.Invoke(this, $"{stepPrefix} ❌ 鼠标滚动失败: {ex.Message}");
                                }
                            }
                        }
                        // 如果鼠标滚动先执行
                        else if (hasMouseScroll && mole.IsMouseScrollExecuteFirst)
                        {
                            var direction = mole.ScrollUp ? "向上" : "向下";
                            LogMessage?.Invoke(this, $"{stepPrefix} 🖱️ 鼠标{direction}滚动 {mole.ScrollCount} 次");
                            
                            try
                            {
                                PerformMouseScroll(mole.ScrollUp, mole.ScrollCount);
                                LogMessage?.Invoke(this, $"{stepPrefix} ✅ 滚动完成");
                                
                                // 滚动后等待
                                if (mole.ScrollWaitMs > 0)
                                {
                                    await Task.Delay(mole.ScrollWaitMs, token);
                                    LogMessage?.Invoke(this, $"{stepPrefix} ⏱️ 滚动后已等待 {mole.ScrollWaitMs}ms");
                                }
                            }
                            catch (Exception ex)
                            {
                                LogMessage?.Invoke(this, $"{stepPrefix} ❌ 鼠标滚动失败: {ex.Message}");
                            }
                            
                            // 然后执行键盘按键
                            if (hasKeyPress)
                            {
                                LogMessage?.Invoke(this, $"{stepPrefix} ⌨️ 发送按键: {mole.KeyPressDefinition}");
                                
                                try
                                {
                                    SendKeyPress(mole.KeyPressDefinition);
                                    LogMessage?.Invoke(this, $"{stepPrefix} ✅ 按键已发送");
                                    
                                    // 等待指定时间
                                    if (mole.KeyPressWaitMs > 0)
                                    {
                                        await Task.Delay(mole.KeyPressWaitMs, token);
                                        LogMessage?.Invoke(this, $"{stepPrefix} ⏱️ 已等待 {mole.KeyPressWaitMs}ms");
                                    }
                                }
                                catch (Exception ex)
                                {
                                    LogMessage?.Invoke(this, $"{stepPrefix} ❌ 按键发送失败: {ex.Message}");
                                }
                            }
                        }
                        // 默认情况：只执行已启用的操作
                        else
                        {
                            if (hasKeyPress)
                            {
                                LogMessage?.Invoke(this, $"{stepPrefix} ⌨️ 发送按键: {mole.KeyPressDefinition}");
                                
                                try
                                {
                                    SendKeyPress(mole.KeyPressDefinition);
                                    LogMessage?.Invoke(this, $"{stepPrefix} ✅ 按键已发送");
                                    
                                    // 等待指定时间
                                    if (mole.KeyPressWaitMs > 0)
                                    {
                                        await Task.Delay(mole.KeyPressWaitMs, token);
                                        LogMessage?.Invoke(this, $"{stepPrefix} ⏱️ 已等待 {mole.KeyPressWaitMs}ms");
                                    }
                                }
                                catch (Exception ex)
                                {
                                    LogMessage?.Invoke(this, $"{stepPrefix} ❌ 按键发送失败: {ex.Message}");
                                }
                            }
                            
                            if (hasMouseScroll)
                            {
                                var direction = mole.ScrollUp ? "向上" : "向下";
                                LogMessage?.Invoke(this, $"{stepPrefix} 🖱️ 鼠标{direction}滚动 {mole.ScrollCount} 次");
                                
                                try
                                {
                                    PerformMouseScroll(mole.ScrollUp, mole.ScrollCount);
                                    LogMessage?.Invoke(this, $"{stepPrefix} ✅ 滚动完成");
                                    
                                    // 滚动后等待
                                    if (mole.ScrollWaitMs > 0)
                                    {
                                        await Task.Delay(mole.ScrollWaitMs, token);
                                        LogMessage?.Invoke(this, $"{stepPrefix} ⏱️ 滚动后已等待 {mole.ScrollWaitMs}ms");
                                    }
                                }
                                catch (Exception ex)
                                {
                                    LogMessage?.Invoke(this, $"{stepPrefix} ❌ 鼠标滚动失败: {ex.Message}");
                                }
                            }
                        }
                    }
                    else
                    {
                        // 跳转模式
                        LogMessage?.Invoke(this, $"{stepPrefix} 🔗 跳转到 {mole.JumpTargetGroup}");
                        
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
                                    
                                    await ExecuteMoleSequenceInternal(targetGroup.Moles, token, totalSteps, targetStartIndex);
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
                        LogMessage?.Invoke(this, $"{stepPrefix} ⏹️ 执行到停止步骤，打地鼠已停止");
                        Stop(); // 停止打地鼠
                        return; // 退出执行
                    }
                    
                    // 执行一次空击
                    ClickAt(mole.IdleClickPosition.Value);
                    LogMessage?.Invoke(this, $"{stepPrefix} 💤 空击 ({mole.IdleClickPosition.Value.X}, {mole.IdleClickPosition.Value.Y})");
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
                        LogMessage?.Invoke(this, $"{stepPrefix} ⏳ 等待图像出现: {mole.Name}");
                        
                        ImageMatchResult? matchResult = null;
                        int waitCount = 0;
                        
                        // 持续扫描直到找到图像
                        while (!token.IsCancellationRequested)
                        {
                            matchResult = FindImageWithEmgu(mole.ImagePath, mole.SimilarityThreshold);
                            
                            if (matchResult != null && matchResult.Found)
                            {
                                // 找到了，退出等待循环
                                break;
                            }
                            
                            waitCount++;
                            if (waitCount % 10 == 0)
                            {
                                LogMessage?.Invoke(this, $"{stepPrefix} ⏳ 继续等待... (已扫描 {waitCount} 次)");
                            }
                            
                            // 等待一小段时间后再次扫描
                            await Task.Delay(100, token);
                        }
                        
                        // 找到后点击
                        if (matchResult != null && matchResult.Found)
                        {
                            ClickAt(matchResult.Center);
                            string scanInfo = waitCount > 0 ? $"（{waitCount}次扫描）" : "";
                            LogMessage?.Invoke(this, $"{stepPrefix} 🎯 {scanInfo}[{mole.Name}] 出现，击中 ({matchResult.Center.X}, {matchResult.Center.Y})");
                            
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
                                        LogMessage?.Invoke(this, $"{stepPrefix} 🔄 持续点击第 {clickCount} 次 ({recheckResult.Center.X}, {recheckResult.Center.Y})");
                                    }
                                    else
                                    {
                                        // 目标已消失，退出循环
                                        LogMessage?.Invoke(this, $"{stepPrefix} ✅ 图像已消失，共点击 {clickCount} 次");
                                        break;
                                    }
                                }
                            }
                            
                            // 如果启用了"点击后等待"
                            if (mole.WaitAfterClick && mole.WaitAfterClickMs > 0)
                            {
                                LogMessage?.Invoke(this, $"{stepPrefix} ⏱️ 等待 {mole.WaitAfterClickMs} ms...");
                                await Task.Delay(mole.WaitAfterClickMs, token);
                            }
                        }
                    }
                    else
                    {
                        // 正常模式：扫描一次（或带超时的等待模式）
                        ImageMatchResult? matchResult = null;
                        int scanCount = 0;
                        
                        // 如果启用了"等待超时后返回上一步"
                        if (mole.ReturnToPreviousOnTimeout && mole.TimeoutMs > 0)
                        {
                            LogMessage?.Invoke(this, $"{stepPrefix} ⏳ 等待图像出现（超时: {mole.TimeoutMs}ms）: {mole.Name}");
                            
                            var startTime = DateTime.Now;
                            
                            // 在超时时间内持续扫描
                            while (!token.IsCancellationRequested)
                            {
                                matchResult = FindImageWithEmgu(mole.ImagePath, mole.SimilarityThreshold);
                                scanCount++;
                                
                                if (matchResult != null && matchResult.Found)
                                {
                                    // 找到了，退出等待循环
                                    break;
                                }
                                
                                // 检查是否超时
                                var elapsed = (DateTime.Now - startTime).TotalMilliseconds;
                                if (elapsed >= mole.TimeoutMs)
                                {
                                    // 超时了，返回上一步
                                    if (i > startIndex)
                                    {
                                        LogMessage?.Invoke(this, $"{stepPrefix} ⏰ 等待超时（{mole.TimeoutMs}ms），返回上一个步骤");
                                        i = i - 2; // -2 是因为循环会 +1，所以实际是回到上一步
                                        currentStep--; // 步骤计数也要回退
                                        await Task.Delay(50, token);
                                        break; // 跳出while循环，继续for循环
                                    }
                                    else
                                    {
                                        // 已经是第一步，无法返回上一步
                                        LogMessage?.Invoke(this, $"{stepPrefix} ⏰ 等待超时（{mole.TimeoutMs}ms），已是第一步，跳过");
                                        matchResult = null; // 确保matchResult为null，后续会跳过
                                        break;
                                    }
                                }
                                
                                // 等待一小段时间后再次扫描
                                await Task.Delay(100, token);
                            }
                            
                            // 如果是因为超时返回上一步，直接continue到下一次循环
                            if (matchResult == null || !matchResult.Found)
                            {
                                await Task.Delay(50, token);
                                continue;
                            }
                        }
                        else
                        {
                            // 正常模式：扫描一次
                            matchResult = FindImageWithEmgu(mole.ImagePath, mole.SimilarityThreshold);
                        }
                        
                        if (matchResult != null && matchResult.Found)
                        {
                            // 找到地鼠，点击中心点
                            ClickAt(matchResult.Center);
                            string scanInfo = scanCount > 0 ? $"（{scanCount}次扫描）" : "";
                            LogMessage?.Invoke(this, $"{stepPrefix} 🎯 {scanInfo}[{mole.Name}] 出现，击中 ({matchResult.Center.X}, {matchResult.Center.Y})");
                            
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
                                        LogMessage?.Invoke(this, $"{stepPrefix} 🔄 持续点击第 {clickCount} 次 ({recheckResult.Center.X}, {recheckResult.Center.Y})");
                                    }
                                    else
                                    {
                                        // 目标已消失，退出循环
                                        LogMessage?.Invoke(this, $"{stepPrefix} ✅ 图像已消失，共点击 {clickCount} 次");
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
                                LogMessage?.Invoke(this, $"{stepPrefix} ⏱️ 等待 {mole.WaitAfterClickMs} ms...");
                                await Task.Delay(mole.WaitAfterClickMs, token);
                            }
                        }
                        else
                        {
                            // 未找到地鼠
                            if (mole.JumpToPreviousOnFail && i > startIndex)
                            {
                                // 启用了"识别失败跳转到上一步"，且不是第一步
                                LogMessage?.Invoke(this, $"{stepPrefix} ⚠️ [{mole.Name}] 未找到，跳转到上一个步骤");
                                i = i - 2; // -2 是因为循环会 +1，所以实际是回到上一步
                                currentStep--; // 步骤计数也要回退
                                await Task.Delay(50, token);
                                continue;
                            }
                            else
                            {
                                // 未找到地鼠，跳过此步骤（默认行为）
                                string confidenceInfo = matchResult != null ? $" (实际匹配 {matchResult.Confidence:F2})" : "";
                                LogMessage?.Invoke(this, $"{stepPrefix} ⏭️ [{mole.Name}] 未找到 (跳过){confidenceInfo}");
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

        /// <summary>
        /// 发送键盘按键
        /// </summary>
        private void SendKeyPress(string keyDefinition)
        {
            if (string.IsNullOrEmpty(keyDefinition))
                return;

            // 解析按键定义（如 "Ctrl+C", "Enter", "F1"）
            var parts = keyDefinition.Split('+');
            var modifiers = new List<Keys>();
            Keys mainKey = Keys.None;

            foreach (var part in parts)
            {
                var trimmedPart = part.Trim();
                
                if (trimmedPart.Equals("Ctrl", StringComparison.OrdinalIgnoreCase))
                {
                    modifiers.Add(Keys.ControlKey);
                }
                else if (trimmedPart.Equals("Shift", StringComparison.OrdinalIgnoreCase))
                {
                    modifiers.Add(Keys.ShiftKey);
                }
                else if (trimmedPart.Equals("Alt", StringComparison.OrdinalIgnoreCase))
                {
                    modifiers.Add(Keys.Menu);
                }
                else
                {
                    // 主键
                    if (Enum.TryParse<Keys>(trimmedPart, true, out var parsedKey))
                    {
                        mainKey = parsedKey;
                    }
                }
            }

            // 按下修饰键
            foreach (var modifier in modifiers)
            {
                WindowHelper.keybd_event((byte)modifier, 0, 0, 0);
                Thread.Sleep(10);
            }

            // 按下主键
            if (mainKey != Keys.None)
            {
                WindowHelper.keybd_event((byte)mainKey, 0, 0, 0);
                Thread.Sleep(10);
                WindowHelper.keybd_event((byte)mainKey, 0, WindowHelper.KEYEVENTF_KEYUP, 0);
            }

            // 释放修饰键（逆序）
            for (int i = modifiers.Count - 1; i >= 0; i--)
            {
                WindowHelper.keybd_event((byte)modifiers[i], 0, WindowHelper.KEYEVENTF_KEYUP, 0);
                Thread.Sleep(10);
            }
        }

        /// <summary>
        /// 执行鼠标滚动操作
        /// </summary>
        /// <param name="scrollUp">true=向上滚动, false=向下滚动</param>
        /// <param name="scrollCount">滚动次数</param>
        private void PerformMouseScroll(bool scrollUp, int scrollCount)
        {
            // 滚动方向：正值向上，负值向下
            int scrollAmount = scrollUp ? WindowHelper.WHEEL_DELTA : -WindowHelper.WHEEL_DELTA;
            
            for (int i = 0; i < scrollCount; i++)
            {
                WindowHelper.mouse_event(WindowHelper.MOUSEEVENTF_WHEEL, 0, 0, (uint)scrollAmount, UIntPtr.Zero);
                Thread.Sleep(50); // 每次滚动之间短暂延迟
            }
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
