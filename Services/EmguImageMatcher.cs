using System;
using System.Collections.Generic;
using System.Drawing;
using Emgu.CV;
using Emgu.CV.CvEnum;
using Emgu.CV.Structure;
using Emgu.CV.Util;

namespace WindowInspector.Services
{
    /// <summary>
    /// 图像匹配结果
    /// </summary>
    public class ImageMatchResult
    {
        public bool Found { get; set; }
        public Point TopLeft { get; set; }
        public Point Center { get; set; }
        public int Width { get; set; }
        public int Height { get; set; }
        public double Confidence { get; set; }
        public string? Error { get; set; }
    }

    /// <summary>
    /// 使用 Emgu.CV (OpenCV C# 封装) 进行图像匹配
    /// 性能比 Python 方案快 5-10 倍
    /// </summary>
    public class EmguImageMatcher : IDisposable
    {
        private bool _disposed = false;
        private readonly Dictionary<string, Mat> _templateCache = new Dictionary<string, Mat>();

        /// <summary>
        /// 在截图中查找模板图像
        /// </summary>
        /// <param name="screenshot">屏幕截图</param>
        /// <param name="templatePath">模板图像路径</param>
        /// <param name="threshold">匹配阈值 (0.0-1.0)</param>
        /// <returns>匹配结果</returns>
        public ImageMatchResult? FindTemplate(Bitmap screenshot, string templatePath, double threshold = 0.85)
        {
            try
            {
                // 从缓存加载模板图像
                var template = LoadTemplateFromCache(templatePath);
                
                if (template == null || template.IsEmpty)
                {
                    return new ImageMatchResult
                    {
                        Found = false,
                        Error = $"无法加载模板图像: {templatePath}"
                    };
                }

                // 将 Bitmap 转换为 Mat
                using var screenshotMat = BitmapToMat(screenshot);
                
                return FindTemplateInternal(screenshotMat, template, threshold);
            }
            catch (Exception ex)
            {
                return new ImageMatchResult
                {
                    Found = false,
                    Error = ex.Message
                };
            }
        }

        /// <summary>
        /// 在截图中查找模板图像（使用 Mat 对象，避免重复转换）
        /// </summary>
        /// <param name="screenshot">屏幕截图 Mat</param>
        /// <param name="templatePath">模板图像路径</param>
        /// <param name="threshold">匹配阈值 (0.0-1.0)</param>
        /// <returns>匹配结果</returns>
        public ImageMatchResult? FindTemplate(Mat screenshot, string templatePath, double threshold = 0.85)
        {
            try
            {
                // 从缓存加载模板图像
                var template = LoadTemplateFromCache(templatePath);
                
                if (template == null || template.IsEmpty)
                {
                    return new ImageMatchResult
                    {
                        Found = false,
                        Error = $"无法加载模板图像: {templatePath}"
                    };
                }

                return FindTemplateInternal(screenshot, template, threshold);
            }
            catch (Exception ex)
            {
                return new ImageMatchResult
                {
                    Found = false,
                    Error = ex.Message
                };
            }
        }

        /// <summary>
        /// 内部匹配方法
        /// </summary>
        private ImageMatchResult FindTemplateInternal(Mat screenshot, Mat template, double threshold)
        {
            try
            {
                // 获取模板尺寸
                int templateWidth = template.Width;
                int templateHeight = template.Height;

                // 执行模板匹配（使用归一化相关系数匹配，效果最好）
                using var result = new Mat();
                CvInvoke.MatchTemplate(screenshot, template, result, TemplateMatchingType.CcoeffNormed);

                // 查找最佳匹配位置
                double minVal = 0, maxVal = 0;
                Point minLoc = new Point(), maxLoc = new Point();
                CvInvoke.MinMaxLoc(result, ref minVal, ref maxVal, ref minLoc, ref maxLoc);

                // 对于 CcoeffNormed，maxVal 是最佳匹配
                double confidence = maxVal;
                Point topLeft = maxLoc;

                // 检查是否达到阈值
                if (confidence < threshold)
                {
                    return new ImageMatchResult
                    {
                        Found = false,
                        Confidence = confidence
                    };
                }

                // 计算中心点
                Point center = new Point(
                    topLeft.X + templateWidth / 2,
                    topLeft.Y + templateHeight / 2
                );

                return new ImageMatchResult
                {
                    Found = true,
                    TopLeft = topLeft,
                    Center = center,
                    Width = templateWidth,
                    Height = templateHeight,
                    Confidence = confidence
                };
            }
            catch (Exception ex)
            {
                return new ImageMatchResult
                {
                    Found = false,
                    Error = ex.Message
                };
            }
        }

        /// <summary>
        /// 将 Bitmap 转换为 Mat
        /// </summary>
        private Mat BitmapToMat(Bitmap bitmap)
        {
            // 将 Bitmap 保存到内存流，然后用 CvInvoke 读取
            using var ms = new System.IO.MemoryStream();
            bitmap.Save(ms, System.Drawing.Imaging.ImageFormat.Bmp);
            ms.Position = 0;
            
            // 从内存流读取为字节数组
            byte[] imageBytes = ms.ToArray();
            
            // 使用 CvInvoke.Imdecode 从字节数组创建 Mat
            using var vectorOfByte = new VectorOfByte(imageBytes);
            Mat result = new Mat();
            CvInvoke.Imdecode(vectorOfByte, ImreadModes.Color, result);
            return result;
        }

        /// <summary>
        /// 从缓存加载模板图像，如果缓存中没有则从磁盘加载并缓存
        /// </summary>
        private Mat? LoadTemplateFromCache(string templatePath)
        {
            // 检查缓存
            if (_templateCache.TryGetValue(templatePath, out var cachedTemplate))
            {
                return cachedTemplate;
            }

            // 从磁盘加载
            var template = CvInvoke.Imread(templatePath, ImreadModes.Color);
            
            if (!template.IsEmpty)
            {
                // 加入缓存
                _templateCache[templatePath] = template;
            }

            return template;
        }

        /// <summary>
        /// 清除模板缓存
        /// </summary>
        public void ClearTemplateCache()
        {
            foreach (var template in _templateCache.Values)
            {
                template?.Dispose();
            }
            _templateCache.Clear();
        }

        /// <summary>
        /// 从 Bitmap 创建 Mat（用于批量匹配，避免重复转换）
        /// </summary>
        public Mat CreateMatFromBitmap(Bitmap bitmap)
        {
            return BitmapToMat(bitmap);
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
                    // 释放托管资源
                    ClearTemplateCache();
                }
                _disposed = true;
            }
        }
    }
}
