using System;
using System.Drawing;
using System.Windows.Forms;

namespace SW2026RibbonAddin.Forms
{
    internal sealed class FastenerReferenceTypeForm : Form
    {
        private readonly RadioButton _rbBolt;
        private readonly RadioButton _rbWasher;
        private readonly RadioButton _rbNut;

        public string SelectedFamily
        {
            get
            {
                if (_rbWasher.Checked) return "Washer";
                if (_rbNut.Checked) return "Nut";
                return "Bolt";
            }
        }

        public FastenerReferenceTypeForm(string initialFamily)
        {
            Text = "Fastener type";
            StartPosition = FormStartPosition.CenterParent;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            ClientSize = new Size(260, 140);

            var lbl = new Label
            {
                Text = "Select the fastener type for this part:",
                AutoSize = false,
                Left = 12,
                Top = 10,
                Width = 230,
                Height = 30
            };
            Controls.Add(lbl);

            _rbBolt = new RadioButton
            {
                Text = "Bolt / Screw",
                Left = 20,
                Top = 40,
                Width = 200
            };
            _rbWasher = new RadioButton
            {
                Text = "Washer",
                Left = 20,
                Top = 60,
                Width = 200
            };
            _rbNut = new RadioButton
            {
                Text = "Nut",
                Left = 20,
                Top = 80,
                Width = 200
            };

            Controls.Add(_rbBolt);
            Controls.Add(_rbWasher);
            Controls.Add(_rbNut);

            var ok = new Button
            {
                Text = "OK",
                DialogResult = DialogResult.OK,
                Left = 70,
                Top = 110,
                Width = 60
            };
            var cancel = new Button
            {
                Text = "Cancel",
                DialogResult = DialogResult.Cancel,
                Left = 140,
                Top = 110,
                Width = 60
            };

            Controls.Add(ok);
            Controls.Add(cancel);

            AcceptButton = ok;
            CancelButton = cancel;

            // Initialize selection
            var fam = initialFamily ?? string.Empty;
            if (fam.Equals("Washer", StringComparison.OrdinalIgnoreCase))
                _rbWasher.Checked = true;
            else if (fam.Equals("Nut", StringComparison.OrdinalIgnoreCase))
                _rbNut.Checked = true;
            else
                _rbBolt.Checked = true;
        }
    }
}
