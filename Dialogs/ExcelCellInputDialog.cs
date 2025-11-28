using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace WindowInspector
{
    public class ExcelCellInputDialog : Form
    {
        private List<TextBox> txtCells = null!;
        private List<Button> btnSelects = null!;
        private Button btnOk = null!;
        private Button btnCancel = null!;
        private Button btnSwitchMode = null!;
        private Panel pnlSingleMode = null!;
        private Panel pnlMultiMode = null!;
        private TextBox txtMultiCells = null!;
        private Button btnMultiSelect = null!;
        private bool isMultiMode = false;
        private int cellCount;

        public List<string> Cells { get; private set; } = new();

        public ExcelCellInputDialog(int cellCount)
        {
            this.cellCount = cellCount;
            InitializeComponent(cellCount);
        }

        private void InitializeComponent(int cellCount)
        {
            Text = "配置Excel单元格地址";
            Size = new Size(500, 200 + cellCount * 40);
            StartPosition = FormStartPosition.CenterParent;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            TopMost = true;  // 强制置顶

            var lblInfo = new Label
            {
                Text = "请输入单元格地址或点击选择按钮在Excel中选择",
                Location = new Point(20, 15),
                Size = new Size(450, 20),
                Parent = this
            };

            // 切换模式按钮
            btnSwitchMode = new Button
            {
                Text = "切换到批量模式",
                Location = new Point(20, 40),
                Size = new Size(120, 25),
                Parent = this
            };
            btnSwitchMode.Click += BtnSwitchMode_Click;

            // 单个模式面板
            pnlSingleMode = new Panel
            {
                Location = new Point(0, 70),
                Size = new Size(500, cellCount * 40 + 10),
                Parent = this
            };

            txtCells = new List<TextBox>();
            btnSelects = new List<Button>();
            
            for (int i = 0; i < cellCount; i++)
            {
                var lbl = new Label
                {
                    Text = $"单元格{i + 1}:",
                    Location = new Point(20, 10 + i * 40),
                    Size = new Size(80, 20),
                    Parent = pnlSingleMode
                };

                var txt = new TextBox
                {
                    Location = new Point(110, 8 + i * 40),
                    Size = new Size(260, 25),
                    Text = $"{(char)('A' + i)}1",
                    Parent = pnlSingleMode
                };
                txtCells.Add(txt);

                var btnSelect = new Button
                {
                    Text = "选择",
                    Location = new Point(380, 7 + i * 40),
                    Size = new Size(80, 27),
                    Tag = i,
                    Parent = pnlSingleMode
                };
                btnSelect.Click += BtnSelect_Click;
                btnSelects.Add(btnSelect);
            }

            // 批量模式面板
            pnlMultiMode = new Panel
            {
                Location = new Point(0, 70),
                Size = new Size(500, 80),
                Visible = false,
                Parent = this
            };

            var lblMulti = new Label
            {
                Text = "单元格地址（可拖选多个）:",
                Location = new Point(20, 15),
                Size = new Size(200, 20),
                Parent = pnlMultiMode
            };

            txtMultiCells = new TextBox
            {
                Location = new Point(20, 40),
                Size = new Size(350, 25),
                Parent = pnlMultiMode
            };

            btnMultiSelect = new Button
            {
                Text = "在Excel中选择",
                Location = new Point(380, 39),
                Size = new Size(100, 27),
                Parent = pnlMultiMode
            };
            btnMultiSelect.Click += BtnMultiSelect_Click;

            // 确定取消按钮（增加50高度）
            var btnY = 130 + (isMultiMode ? 80 : cellCount * 40 + 10);

            btnOk = new Button
            {
                Text = "确定",
                Location = new Point(290, btnY),
                Size = new Size(80, 30),
                DialogResult = DialogResult.OK,
                Parent = this
            };
            btnOk.Click += BtnOk_Click;

            btnCancel = new Button
            {
                Text = "取消",
                Location = new Point(380, btnY),
                Size = new Size(80, 30),
                DialogResult = DialogResult.Cancel,
                Parent = this
            };

            AcceptButton = btnOk;
            CancelButton = btnCancel;
        }

        private void BtnSwitchMode_Click(object? sender, EventArgs e)
        {
            isMultiMode = !isMultiMode;
            
            if (isMultiMode)
            {
                pnlSingleMode.Visible = false;
                pnlMultiMode.Visible = true;
                btnSwitchMode.Text = "切换到单个模式";
                
                var cells = txtCells.Select(t => t.Text.Trim()).Where(s => !string.IsNullOrEmpty(s));
                txtMultiCells.Text = string.Join(",", cells);
                
                Height = 200 + 80;
            }
            else
            {
                pnlSingleMode.Visible = true;
                pnlMultiMode.Visible = false;
                btnSwitchMode.Text = "切换到批量模式";
                
                var cells = ParseMultiCells(txtMultiCells.Text);
                for (int i = 0; i < Math.Min(cells.Count, txtCells.Count); i++)
                {
                    txtCells[i].Text = cells[i];
                }
                
                Height = 200 + cellCount * 40;
            }
            
            var btnY = 130 + (isMultiMode ? 80 : cellCount * 40 + 10);
            btnOk.Location = new Point(290, btnY);
            btnCancel.Location = new Point(380, btnY);
        }

        private void BtnSelect_Click(object? sender, EventArgs e)
        {
            var btn = sender as Button;
            if (btn == null) return;
            
            int index = (int)btn.Tag!;
            var cell = GetSelectedCellFromExcel();
            
            if (!string.IsNullOrEmpty(cell))
            {
                txtCells[index].Text = cell;
            }
        }

        private void BtnMultiSelect_Click(object? sender, EventArgs e)
        {
            var cells = GetSelectedCellsFromExcel();
            if (cells.Count > 0)
            {
                txtMultiCells.Text = string.Join(",", cells);
            }
        }

        [DllImport("oleaut32.dll", PreserveSig = false)]
        private static extern void GetActiveObject(ref Guid rclsid, IntPtr pvReserved, [MarshalAs(UnmanagedType.IUnknown)] out object ppunk);

        private string GetSelectedCellFromExcel()
        {
            try
            {
                var clsid = new Guid("00024500-0000-0000-C000-000000000046");
                GetActiveObject(ref clsid, IntPtr.Zero, out object excelObj);
                dynamic excel = excelObj;
                dynamic selection = excel.Selection;
                
                if (selection != null)
                {
                    string address = selection.Address;
                    address = address.Replace("$", "");
                    return address;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"获取Excel选中单元格失败: {ex.Message}\n\n请确保Excel已打开并选中了单元格", 
                    "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            
            return string.Empty;
        }

        private List<string> GetSelectedCellsFromExcel()
        {
            var cells = new List<string>();
            
            try
            {
                var clsid = new Guid("00024500-0000-0000-C000-000000000046");
                GetActiveObject(ref clsid, IntPtr.Zero, out object excelObj);
                dynamic excel = excelObj;
                dynamic selection = excel.Selection;
                
                if (selection != null)
                {
                    string address = selection.Address;
                    address = address.Replace("$", "");
                    
                    if (address.Contains(":"))
                    {
                        var parts = address.Split(':');
                        if (parts.Length == 2)
                        {
                            cells = ExpandRange(parts[0], parts[1]);
                        }
                    }
                    else
                    {
                        cells.Add(address);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"获取Excel选中单元格失败: {ex.Message}\n\n请确保Excel已打开并选中了单元格", 
                    "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            
            return cells;
        }

        private List<string> ExpandRange(string start, string end)
        {
            var cells = new List<string>();
            
            try
            {
                var startCol = new string(start.TakeWhile(char.IsLetter).ToArray());
                var startRow = int.Parse(new string(start.SkipWhile(char.IsLetter).ToArray()));
                var endCol = new string(end.TakeWhile(char.IsLetter).ToArray());
                var endRow = int.Parse(new string(end.SkipWhile(char.IsLetter).ToArray()));
                
                int startColNum = ColumnLetterToNumber(startCol);
                int endColNum = ColumnLetterToNumber(endCol);
                
                for (int row = startRow; row <= endRow; row++)
                {
                    for (int col = startColNum; col <= endColNum; col++)
                    {
                        cells.Add(ColumnNumberToLetter(col) + row);
                    }
                }
            }
            catch { }
            
            return cells;
        }

        private int ColumnLetterToNumber(string column)
        {
            int result = 0;
            for (int i = 0; i < column.Length; i++)
            {
                result = result * 26 + (column[i] - 'A' + 1);
            }
            return result;
        }

        private string ColumnNumberToLetter(int column)
        {
            string result = "";
            while (column > 0)
            {
                int modulo = (column - 1) % 26;
                result = (char)('A' + modulo) + result;
                column = (column - modulo) / 26;
            }
            return result;
        }

        private List<string> ParseMultiCells(string input)
        {
            var cells = new List<string>();
            
            if (string.IsNullOrWhiteSpace(input))
                return cells;
            
            var parts = input.Split(new[] { ',', ' ', ';' }, StringSplitOptions.RemoveEmptyEntries);
            
            foreach (var part in parts)
            {
                var trimmed = part.Trim().ToUpper();
                
                if (trimmed.Contains(":"))
                {
                    var rangeParts = trimmed.Split(':');
                    if (rangeParts.Length == 2)
                    {
                        cells.AddRange(ExpandRange(rangeParts[0], rangeParts[1]));
                    }
                }
                else
                {
                    cells.Add(trimmed);
                }
            }
            
            return cells;
        }

        private void BtnOk_Click(object? sender, EventArgs e)
        {
            Cells = new List<string>();
            
            if (isMultiMode)
            {
                Cells = ParseMultiCells(txtMultiCells.Text);
                
                if (Cells.Count == 0)
                {
                    MessageBox.Show("请输入或选择单元格地址", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    DialogResult = DialogResult.None;
                    return;
                }
            }
            else
            {
                foreach (var txt in txtCells)
                {
                    if (string.IsNullOrWhiteSpace(txt.Text))
                    {
                        MessageBox.Show("请填写所有单元格地址", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        DialogResult = DialogResult.None;
                        return;
                    }
                    Cells.Add(txt.Text.Trim().ToUpper());
                }
            }
        }
    }
}
