using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;

namespace WindowInspector
{
    public class TextInputDialog : Form
    {
        private TextBox txtName = null!;
        private List<TextBox> txtInputs = null!;
        private Button btnOk = null!;
        private Button btnCancel = null!;

        public string ItemName { get; private set; } = string.Empty;
        public List<string> Texts { get; private set; } = new();

        public TextInputDialog(int inputCount)
        {
            InitializeDialog(inputCount);
        }

        private void InitializeDialog(int inputCount)
        {
            Text = "保存输入内容";
            Size = new Size(400, 150 + inputCount * 40);
            StartPosition = FormStartPosition.CenterParent;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;

            var lblName = new Label
            {
                Text = "名称:",
                Location = new Point(20, 20),
                Size = new Size(60, 20),
                Parent = this
            };

            txtName = new TextBox
            {
                Location = new Point(90, 18),
                Size = new Size(280, 25),
                Parent = this
            };

            txtInputs = new List<TextBox>();
            for (int i = 0; i < inputCount; i++)
            {
                var lbl = new Label
                {
                    Text = $"文本{i + 1}:",
                    Location = new Point(20, 60 + i * 40),
                    Size = new Size(60, 20),
                    Parent = this
                };

                var txt = new TextBox
                {
                    Location = new Point(90, 58 + i * 40),
                    Size = new Size(280, 25),
                    Parent = this
                };
                
                // 如果只有2个文本且是第2个，使用密码框
                if (inputCount == 2 && i == 1)
                {
                    txt.UseSystemPasswordChar = true;
                }
                
                txtInputs.Add(txt);
            }

            var btnY = 70 + inputCount * 40;

            btnOk = new Button
            {
                Text = "确定",
                Location = new Point(200, btnY),
                Size = new Size(80, 30),
                DialogResult = DialogResult.OK,
                Parent = this
            };
            btnOk.Click += BtnOk_Click;

            btnCancel = new Button
            {
                Text = "取消",
                Location = new Point(290, btnY),
                Size = new Size(80, 30),
                DialogResult = DialogResult.Cancel,
                Parent = this
            };

            AcceptButton = btnOk;
            CancelButton = btnCancel;
        }

        private void BtnOk_Click(object? sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(txtName.Text))
            {
                MessageBox.Show("请输入名称", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                DialogResult = DialogResult.None;
                return;
            }

            ItemName = txtName.Text.Trim();
            Texts = new List<string>();
            foreach (var txt in txtInputs)
            {
                Texts.Add(txt.Text);
            }
        }
    }
}
