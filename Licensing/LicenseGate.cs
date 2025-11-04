// File: Licensing/LicenseGate.cs
using Microsoft.Win32;
using System;
using System.Windows.Forms;
using Licensing;

namespace SW2025RibbonAddin
{
    internal static class LicenseGate
    {
        // Must match the "prd" claim you issue from Key Ring
        private const string ExpectedProduct = "SW2025RibbonAddin";

        // Paste one JWK or a JWKS {"keys":[...]} exported from Key Ring (PUBLIC KEYS ONLY)
        private const string JwkOrJwksJson = @"
        {
          ""kty"": ""EC"",
          ""crv"": ""P-256"",
          ""alg"": ""ES256"",
          ""x"": ""<PASTE X FROM KEY RING>"",
          ""y"": ""<PASTE Y FROM KEY RING>"",
          ""use"": ""sig"",
          ""kid"": ""<PASTE KID FROM KEY RING>""
        }";

        private static Es256LicenseVerifier _verifier;

        private static Es256LicenseVerifier Verifier
        {
            get
            {
                if (_verifier != null) return _verifier;

                var trust = LicenseTrustStore.FromJson(JwkOrJwksJson);
                _verifier = new Es256LicenseVerifier(
                    trust,
                    ExpectedProduct,
                    GetMachineId,
                    TimeSpan.FromMinutes(5));
                return _verifier;
            }
        }

        internal static bool EnsureLicensed(IWin32Window parent = null)
        {
            var token = LoadToken();
            var checkedSaved = false;

            if (!string.IsNullOrWhiteSpace(token))
            {
                var res = Verifier.Verify(token);
                if (res.Success) return true;

                checkedSaved = true;
                MessageBox.Show(parent, "Saved license is not valid:\r\n" + res.Error,
                    "License", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }

            var pasted = PromptForToken(parent);
            if (string.IsNullOrWhiteSpace(pasted))
                return false;

            var res2 = Verifier.Verify(pasted.Trim());
            if (!res2.Success)
            {
                MessageBox.Show(parent, "License is invalid:\r\n" + res2.Error,
                    "License", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }

            SaveToken(pasted.Trim());
            if (checkedSaved) MessageBox.Show(parent, "License activated successfully.",
                "License", MessageBoxButtons.OK, MessageBoxIcon.Information);

            return true;
        }

        private const string RegPath = @"Software\Mehdi\SW2025RibbonAddin";
        private const string RegName = "LicenseToken";

        private static string LoadToken()
        {
            try
            {
                using (var k = Registry.CurrentUser.OpenSubKey(RegPath))
                    return (string)k.GetValue(RegName, "");
            }
            catch { return ""; }
        }

        private static void SaveToken(string token)
        {
            try
            {
                using (var k = Registry.CurrentUser.CreateSubKey(RegPath))
                {
                    if (k != null) k.SetValue(RegName, token, RegistryValueKind.String);
                }
            }
            catch { }
        }

        public static string GetMachineId()
        {
            try
            {
                using (var k = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Microsoft\Cryptography"))
                {
                    var v = (string)k.GetValue("MachineGuid");
                    if (!string.IsNullOrWhiteSpace(v)) return v.Trim().ToUpperInvariant();
                }
            }
            catch { }
            return Environment.MachineName.ToUpperInvariant();
        }

        private static string PromptForToken(IWin32Window parent)
        {
            using (var dlg = new Form
            {
                Text = "Activate License",
                StartPosition = FormStartPosition.CenterParent,
                Width = 640,
                Height = 200,
                MinimizeBox = false,
                MaximizeBox = false,
                ShowInTaskbar = false,
                TopMost = true
            })
            {
                var lbl = new Label
                {
                    Left = 12,
                    Top = 12,
                    Width = 600,
                    Text = "Paste your license token (compact JWS):"
                };
                var box = new TextBox
                {
                    Left = 12,
                    Top = 36,
                    Width = 600,
                    Height = 70,
                    Multiline = true,
                    ScrollBars = ScrollBars.Vertical,
                    AcceptsReturn = true
                };
                var ok = new Button { Text = "Activate", DialogResult = DialogResult.OK, Left = 412, Width = 90, Top = 120 };
                var cancel = new Button { Text = "Cancel", DialogResult = DialogResult.Cancel, Left = 522, Width = 90, Top = 120 };

                dlg.Controls.AddRange(new Control[] { lbl, box, ok, cancel });
                dlg.AcceptButton = ok; dlg.CancelButton = cancel;

                return dlg.ShowDialog(parent) == DialogResult.OK ? box.Text : null;
            }
        }
    }
}
