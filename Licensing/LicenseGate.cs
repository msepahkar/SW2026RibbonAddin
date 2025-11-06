using Microsoft.Win32;
using System;
using System.Windows.Forms;

namespace SW2025RibbonAddin.Licensing
{
    public static class LicenseGate
    {
        public static string MachineId
        {
            get
            {
                try
                {
                    using (var key = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Microsoft\Cryptography"))
                    {
                        var v = key != null ? key.GetValue("MachineGuid") as string : null;
                        if (!string.IsNullOrWhiteSpace(v)) return v.Trim();
                    }
                }
                catch { }
                return (Environment.MachineName ?? "UNKNOWN").ToUpperInvariant();
            }
        }

        public static bool IsLicensed
        {
            get { VerifiedLicense lic; string why; return IsActivated(out lic, out why); }
        }

        public static bool IsActivated(out VerifiedLicense license, out string reason)
        {
            license = null; reason = null;

            string vfErr; LicenseVerifier verifier;
            if (!TryBuildVerifier(out verifier, out vfErr))
            {
                reason = vfErr; return false;
            }

            var token = CleanToken(LoadToken());
            if (string.IsNullOrWhiteSpace(token))
            {
                reason = "No token stored."; return false;
            }

            string err;
            if (!verifier.TryVerify(token, out license, out err))
            {
                reason = err ?? "Verification failed."; return false;
            }
            return true;
        }

        public static bool Activate(string userName, string token, out string error)
        {
            error = null;

            string vfErr; LicenseVerifier verifier;
            if (!TryBuildVerifier(out verifier, out vfErr))
            {
                error = vfErr; return false;
            }

            token = CleanToken(token);

            VerifiedLicense lic;
            if (!verifier.TryVerify(token, out lic, out error))
                return false;

            Save(LicenseSettings.RegistryValueUser, (userName ?? "").Trim());
            Save(LicenseSettings.RegistryValueToken, token);
            return true;
        }

        public static void Deactivate()
        {
            try
            {
                using (var key = Registry.CurrentUser.CreateSubKey(LicenseSettings.RegistryPath, true))
                {
                    if (key != null) key.DeleteValue(LicenseSettings.RegistryValueToken, false);
                }
            }
            catch { }
        }

        public static string LoadStoredUserName()
        {
            using (var key = Registry.CurrentUser.OpenSubKey(LicenseSettings.RegistryPath))
                return key != null ? (key.GetValue(LicenseSettings.RegistryValueUser) as string ?? "") : "";
        }

        // shims if older code calls these
        public static bool TryLoadSavedLicense() { VerifiedLicense lic; string why; return IsActivated(out lic, out why); }
        public static bool TryLoadSavedLicense(out VerifiedLicense license) { string why; return IsActivated(out license, out why); }
        public static bool TryLoadSavedLicense(out string reason) { VerifiedLicense lic; return IsActivated(out lic, out reason); }
        public static void ShowRegistrationDialog() { LicensingUI.ShowRegistrationDialog(null); }
        public static void ShowRegistrationDialog(IWin32Window owner) { LicensingUI.ShowRegistrationDialog(owner); }

        // internals
        private static bool TryBuildVerifier(out LicenseVerifier verifier, out string error)
        {
            verifier = null; error = null;

            var json = LicenseSettings.TrustedKeysJson ?? "";
            if (string.IsNullOrWhiteSpace(json))
            {
                error = "Trusted public key is not configured.";
                return false;
            }

            try
            {
                verifier = new LicenseVerifier(
                    LicenseSettings.TrustedKeysJson,
                    LicenseSettings.Product,
                    MachineId,
                    LicenseSettings.AllowedClockSkew,
                    LicenseSettings.ExpiryGrace);
                return true;
            }
            catch (Exception ex)
            {
                error = ex.Message; return false;
            }
        }

        private static string CleanToken(string token)
        {
            if (string.IsNullOrEmpty(token)) return "";
            token = token.Trim().Replace(" ", "").Replace("\r", "").Replace("\n", "");
            while (token.StartsWith(".")) token = token.Substring(1);
            while (token.EndsWith(".")) token = token.Substring(0, token.Length - 1);
            return token;
        }

        private static string LoadToken()
        {
            using (var key = Registry.CurrentUser.OpenSubKey(LicenseSettings.RegistryPath))
                return key != null ? (key.GetValue(LicenseSettings.RegistryValueToken) as string ?? "") : "";
        }

        private static void Save(string name, string value)
        {
            using (var key = Registry.CurrentUser.CreateSubKey(LicenseSettings.RegistryPath, true))
            {
                if (key != null) key.SetValue(name, value, RegistryValueKind.String);
            }
        }
    }
}
