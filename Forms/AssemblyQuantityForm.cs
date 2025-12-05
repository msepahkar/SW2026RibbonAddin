using System;
using System.Drawing;
using System.Windows.Forms;

namespace SW2026RibbonAddin.Forms
{
    internal sealed class AssemblyQuantityForm : Form
    {
        private readonly NumericUpDown _numQty;
        private readonly Button _okButton;
        private readonly Button _cancelButton;

        public int AssemblyQuantity => (int)_numQty.Value;

        public AssemblyQuantityForm(int initialValue)
        {
            Text = "Number of assemblies";
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            StartPosition = FormStartPosition.CenterParent;
            ClientSize = new Size(260, 110);
            ShowInTaskbar = false;

            var label = new Label
            {
                Left = 12,
                Top = 12,
                Width = 230,
                Text = "Enter number of assemblies:"
            };
            Controls.Add(label);

            _numQty = new NumericUpDown
            {
                Left = 12,
                Top = 35,
                Width = 80,
                Minimum = 1,
                Maximum = 100000,
                Value = Math.Max(1, Math.Min(initialValue, 100000))
            };
            Controls.Add(_numQty);

            _okButton = new Button
            {
                Text = "OK",
                DialogResult = DialogResult.OK,
                Left = 110,
                Top = 70,
                Width = 60
            };
            Controls.Add(_okButton);

            _cancelButton = new Button
            {
                Text = "Cancel",
                DialogResult = DialogResult.Cancel,
                Left = 180,
                Top = 70,
                Width = 60
            };
            Controls.Add(_cancelButton);

            AcceptButton = _okButton;
            CancelButton = _cancelButton;
        }
    }
}
