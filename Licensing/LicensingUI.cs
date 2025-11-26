using System;
using System.Drawing;
using System.Windows.Forms;

namespace SW2026RibbonAddin.Licensing
{
    public static class LicensingUI
    {
        public static void ShowRegistrationDialog(IWin32Window owner = null)
        {
            using (var dlg = new Form())
            {
                dlg.Text = "Register " + LicenseSettings.Product;
                dlg.StartPosition = FormStartPosition.CenterParent;
                dlg.FormBorderStyle = FormBorderStyle.FixedDialog;
                dlg.MaximizeBox = false;
                dlg.MinimizeBox = false;
                dlg.ClientSize = new Size(700, 460);

                var lblProduct = new Label { Text = "Product:", AutoSize = true, Location = new Point(12, 15) };
                var txtProduct = new TextBox { ReadOnly = true, Location = new Point(100, 12), Width = 560, Text = LicenseSettings.Product };

                var lblMachine = new Label { Text = "Machine Id:", AutoSize = true, Location = new Point(12, 50) };
                var txtMachine = new TextBox { ReadOnly = true, Location = new Point(100, 47), Width = 560, Text = LicenseGate.MachineId };

                var lblUser = new Label { Text = "User name:", AutoSize = true, Location = new Point(12, 85) };
                var txtUser = new TextBox { ReadOnly = false, Location = new Point(100, 82), Width = 560, Text = LicenseGate.LoadStoredUserName() };

                var lblToken = new Label { Text = "License token:", AutoSize = true, Location = new Point(12, 120) };

                VerifiedLicense lic; string why;
                var lblStatus = new Label { AutoSize = true, Location = new Point(100, 120) };
                lblStatus.Text = LicenseGate.IsActivated(out lic, out why) ? "Status: Activated" : "Status: Not activated";

                var txtToken = new TextBox
                {
                    Location = new Point(15, 145),
                    Size = new Size(645, 220),
                    Multiline = true,
                    ScrollBars = ScrollBars.Vertical,
                    WordWrap = false
                };

                var btnCopyReq = new Button { Text = "Copy Request", Location = new Point(15, 380), Width = 120 };
                var btnClose = new Button { Text = "Close", Location = new Point(430, 380), Width = 75 };
                var btnActivate = new Button { Text = "Activate", Location = new Point(510, 380), Width = 75 };
                var btnDeactivate = new Button { Text = "Deactivate", Location = new Point(600, 380), Width = 75 };

                btnCopyReq.Click += (s, e) =>
                {
                    var req = "{\"prd\":\"" + LicenseSettings.Product + "\",\"mid\":\"" + LicenseGate.MachineId + "\",\"usr\":\"" + (txtUser.Text ?? "").Trim() + "\"}";
                    Clipboard.SetText(req);
                    MessageBox.Show(dlg, "License request copied to clipboard.", dlg.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
                };

                btnActivate.Click += (s, e) =>
                {
                    var token = (txtToken.Text ?? "");
                    token = token.Trim().Replace(" ", "").Replace("\r", "").Replace("\n", "");
                    while (token.StartsWith(".")) token = token.Substring(1);
                    while (token.EndsWith(".")) token = token.Substring(0, token.Length - 1);
                    txtToken.Text = token;

                    var user = (txtUser.Text ?? "").Trim();
                    if (user.Length == 0) { MessageBox.Show(dlg, "Please enter a user name.", dlg.Text, MessageBoxButtons.OK, MessageBoxIcon.Warning); return; }
                    if (token.Length == 0) { MessageBox.Show(dlg, "Please paste a license token.", dlg.Text, MessageBoxButtons.OK, MessageBoxIcon.Warning); return; }

                    string err;
                    if (!LicenseGate.Activate(user, token, out err))
                    {
                        MessageBox.Show(dlg, "Activation failed: " + err, dlg.Text, MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }

                    MessageBox.Show(dlg, "Activated successfully.", dlg.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
                    lblStatus.Text = "Status: Activated";
                    dlg.DialogResult = DialogResult.OK;
                    dlg.Close();
                };

                btnDeactivate.Click += (s, e) =>
                {
                    var confirm = MessageBox.Show(
                        dlg,
                        "Deactivate and remove the saved license token for this user?\nYou can activate again later with a valid token.",
                        dlg.Text,
                        MessageBoxButtons.YesNo,
                        MessageBoxIcon.Question);

                    if (confirm == DialogResult.Yes)
                    {
                        LicenseGate.Deactivate();
                        MessageBox.Show(dlg, "License removed.", dlg.Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
                        lblStatus.Text = "Status: Not activated";
                    }
                };

                btnClose.Click += (s, e) => dlg.Close();

                dlg.AcceptButton = btnActivate;
                dlg.CancelButton = btnClose; // <-- Close does nothing (safe)

                dlg.Controls.AddRange(new Control[]
                {
                    lblProduct, txtProduct,
                    lblMachine, txtMachine,
                    lblUser, txtUser,
                    lblToken, lblStatus, txtToken,
                    btnCopyReq, btnClose, btnActivate, btnDeactivate
                });

                dlg.ShowDialog(owner);
            }
        }
    }
}
