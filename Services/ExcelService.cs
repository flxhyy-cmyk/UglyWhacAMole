using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using WindowInspector.Models;

namespace WindowInspector.Services
{
    public class ExcelService
    {
        private static bool _excelAvailable = true;

        private static bool IsExcelAvailable()
        {
            if (!_excelAvailable)
                return false;

            try
            {
                var type = Type.GetTypeFromProgID("Excel.Application");
                return type != null;
            }
            catch
            {
                _excelAvailable = false;
                return false;
            }
        }

        public List<SavedTextItem> LoadFromExcel(string filePath, List<string> cells)
        {
            if (!IsExcelAvailable())
            {
                throw new Exception("未检测到Microsoft Excel，请安装Office后重试");
            }

            var items = new List<SavedTextItem>();
            dynamic? excelApp = null;
            dynamic? workbook = null;

            try
            {
                var excelType = Type.GetTypeFromProgID("Excel.Application");
                if (excelType == null)
                    throw new Exception("无法创建Excel应用程序");

                excelApp = Activator.CreateInstance(excelType);
                workbook = excelApp.Workbooks.Open(filePath);
                dynamic worksheet = workbook.ActiveSheet;

                dynamic usedRange = worksheet.UsedRange;
                int rowCount = usedRange.Rows.Count;

                // 检查第一行是否为表头（包含"名称"列）
                dynamic firstCellRange = worksheet.Range["A1"];
                string firstCellValue = firstCellRange?.Value?.ToString() ?? string.Empty;
                bool hasHeader = firstCellValue == "名称";

                int startRow = hasHeader ? 2 : 1; // 如果有表头，从第2行开始

                for (int row = startRow; row <= rowCount; row++)
                {
                    var texts = new List<string>();
                    var hasData = false;
                    string itemName = string.Empty;

                    // 如果有表头格式，读取名称列（A列）
                    if (hasHeader)
                    {
                        dynamic nameRange = worksheet.Range["A" + row];
                        itemName = nameRange?.Value?.ToString() ?? $"Excel行{row}";
                        
                        // 读取文本列（从B列开始）
                        for (int i = 0; i < cells.Count; i++)
                        {
                            var colLetter = GetColumnLetter(i + 1); // B, C, D...
                            dynamic range = worksheet.Range[colLetter + row];
                            var value = range?.Value?.ToString() ?? string.Empty;
                            texts.Add(value);
                            if (!string.IsNullOrWhiteSpace(value))
                                hasData = true;
                        }
                    }
                    else
                    {
                        // 旧格式：直接使用指定的单元格地址
                        itemName = $"Excel行{row}";
                        foreach (var cell in cells)
                        {
                            dynamic range = worksheet.Range[cell + row];
                            var value = range?.Value?.ToString() ?? string.Empty;
                            texts.Add(value);
                            if (!string.IsNullOrWhiteSpace(value))
                                hasData = true;
                        }
                    }

                    if (hasData)
                    {
                        items.Add(new SavedTextItem
                        {
                            Name = itemName,
                            Texts = texts,
                            FromExcel = true,
                            LastFilledTime = null
                        });
                    }
                }
            }
            catch (Exception ex)
            {
                throw new Exception($"加载Excel失败: {ex.Message}");
            }
            finally
            {
                try
                {
                    workbook?.Close(false);
                    excelApp?.Quit();
                    if (workbook != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);
                    if (excelApp != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
                }
                catch { }
            }

            return items;
        }

        public void ExportToExcel(string filePath, List<SavedTextItem> items, List<string> cells)
        {
            if (!IsExcelAvailable())
            {
                throw new Exception("未检测到Microsoft Excel，请安装Office后重试");
            }

            dynamic? excelApp = null;
            dynamic? workbook = null;

            try
            {
                var excelType = Type.GetTypeFromProgID("Excel.Application");
                if (excelType == null)
                    throw new Exception("无法创建Excel应用程序");

                excelApp = Activator.CreateInstance(excelType);
                workbook = excelApp.Workbooks.Add();
                dynamic worksheet = workbook.ActiveSheet;

                // 写入表头
                dynamic headerRange = worksheet.Range["A1"];
                if (headerRange != null)
                    headerRange.Value = "名称";

                // 确定文本列数量
                int textColumnCount = items.Count > 0 ? items[0].Texts.Count : cells.Count;
                
                for (int j = 0; j < textColumnCount; j++)
                {
                    var colLetter = GetColumnLetter(j + 2); // B, C, D... (从第2列开始)
                    dynamic range = worksheet.Range[colLetter + "1"];
                    if (range != null)
                        range.Value = $"文本{j + 1}";
                }

                // 设置表头样式（A列到最后一个文本列）
                var lastColLetter = GetColumnLetter(textColumnCount + 1);
                dynamic headerRow = worksheet.Range["A1:" + lastColLetter + "1"];
                if (headerRow != null)
                {
                    headerRow.Font.Bold = true;
                    headerRow.Interior.Color = 0xD3D3D3; // 浅灰色背景
                }

                // 写入数据（从第2行开始）
                for (int i = 0; i < items.Count; i++)
                {
                    var item = items[i];
                    var rowNum = i + 2; // 从第2行开始（第1行是表头）

                    // 写入名称
                    dynamic nameRange = worksheet.Range["A" + rowNum];
                    if (nameRange != null)
                        nameRange.Value = item.Name;

                    // 写入文本内容
                    for (int j = 0; j < item.Texts.Count; j++)
                    {
                        var colLetter = GetColumnLetter(j + 2); // B, C, D... (从第2列开始)
                        dynamic range = worksheet.Range[colLetter + rowNum];
                        if (range != null)
                            range.Value = item.Texts[j];
                    }
                }

                // 自动调整列宽
                dynamic usedRange = worksheet.UsedRange;
                if (usedRange != null)
                    usedRange.Columns.AutoFit();

                workbook.SaveAs(filePath);
            }
            catch (Exception ex)
            {
                throw new Exception($"导出Excel失败: {ex.Message}");
            }
            finally
            {
                try
                {
                    workbook?.Close(true);
                    excelApp?.Quit();
                    if (workbook != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);
                    if (excelApp != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
                }
                catch { }
            }
        }

        private string GetColumnLetter(int columnNumber)
        {
            string columnLetter = "";
            while (columnNumber > 0)
            {
                int modulo = (columnNumber - 1) % 26;
                columnLetter = Convert.ToChar('A' + modulo) + columnLetter;
                columnNumber = (columnNumber - modulo) / 26;
            }
            return columnLetter;
        }

        public List<SavedTextItem> LoadFromExcelAuto(string filePath)
        {
            if (!IsExcelAvailable())
            {
                throw new Exception("未检测到Microsoft Excel，请安装Office后重试");
            }

            var items = new List<SavedTextItem>();
            dynamic? excelApp = null;
            dynamic? workbook = null;

            try
            {
                var excelType = Type.GetTypeFromProgID("Excel.Application");
                if (excelType == null)
                    throw new Exception("无法创建Excel应用程序");

                excelApp = Activator.CreateInstance(excelType);
                workbook = excelApp.Workbooks.Open(filePath);
                dynamic worksheet = workbook.ActiveSheet;

                dynamic usedRange = worksheet.UsedRange;
                int rowCount = usedRange.Rows.Count;
                int colCount = usedRange.Columns.Count;

                if (rowCount < 2)
                {
                    return items; // 没有数据行
                }

                // 检查第一行第一列是否为"名称"，判断是否有表头
                dynamic firstCellRange = worksheet.Range["A1"];
                string firstCellValue = firstCellRange?.Value?.ToString() ?? string.Empty;
                bool hasHeader = firstCellValue == "名称";

                int startRow = hasHeader ? 2 : 1; // 如果有表头，从第2行开始

                // 统一逻辑：A列是名称，B列开始是文本
                for (int row = startRow; row <= rowCount; row++)
                {
                    // 读取名称（A列）
                    dynamic nameRange = worksheet.Range["A" + row];
                    string itemName = nameRange?.Value?.ToString() ?? string.Empty;

                    if (string.IsNullOrWhiteSpace(itemName))
                        continue; // 跳过空行

                    // 读取所有文本列（从B列开始）
                    var texts = new List<string>();
                    bool hasData = false;

                    for (int col = 2; col <= colCount; col++)
                    {
                        var colLetter = GetColumnLetter(col);
                        dynamic range = worksheet.Range[colLetter + row];
                        var value = range?.Value?.ToString() ?? string.Empty;
                        texts.Add(value);
                        if (!string.IsNullOrWhiteSpace(value))
                            hasData = true;
                    }

                    if (hasData)
                    {
                        items.Add(new SavedTextItem
                        {
                            Name = itemName,
                            Texts = texts,
                            FromExcel = true,
                            LastFilledTime = null
                        });
                    }
                }
            }
            catch (Exception ex)
            {
                throw new Exception($"加载Excel失败: {ex.Message}");
            }
            finally
            {
                try
                {
                    workbook?.Close(false);
                    excelApp?.Quit();
                    if (workbook != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);
                    if (excelApp != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
                }
                catch { }
            }

            return items;
        }

        public void OpenExcel(string filePath)
        {
            try
            {
                // 使用系统默认程序打开Excel文件
                var psi = new ProcessStartInfo
                {
                    FileName = filePath,
                    UseShellExecute = true
                };
                Process.Start(psi);
            }
            catch (Exception ex)
            {
                throw new Exception($"打开Excel失败: {ex.Message}");
            }
        }
    }
}
