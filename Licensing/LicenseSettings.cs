using System;

namespace SW2025RibbonAddin.Licensing
{
    /// <summary>Developer-supplied settings (single source of truth).</summary>
    public static class LicenseSettings
    {
        /// <summary>Product code that must match the "prd" claim.</summary>
        public const string Product = "SW2025RibbonAddin";

        /// <summary>
        /// Trusted public key(s) in JWK or JWKS JSON. Your current P-256 public key:
        /// </summary>
        public const string TrustedKeysJson = @"{
  ""keys"": [
    {
      ""kty"": ""EC"",
      ""crv"": ""P-256"",
      ""alg"": ""ES256"",
      ""x"": ""ZJKssGhuYhXtNLadXPJ47Q3YSxnibrO5pXtsem6esFQ"",
      ""y"": ""nvFdRAmafamXQdOtPZ0WtIj6-QR4jvFXLR-eh7G8BqM"",
      ""use"": ""sig"",
      ""kid"": ""D9LqRzkty2oJTGv1EiOkakmY4IY""
    }
  ]
}";

        /// <summary>Allowed clock skew for iat/nbf checks.</summary>
        public static readonly TimeSpan AllowedClockSkew = TimeSpan.FromMinutes(5);

        /// <summary>Post-expiration grace period.</summary>
        public static readonly TimeSpan ExpiryGrace = TimeSpan.Zero;

        /// <summary>Registry location (HKCU) where token and user name are stored.</summary>
        public const string RegistryPath = @"Software\Mehdi\SW2025RibbonAddin";
        public const string RegistryValueToken = "LicenseToken";
        public const string RegistryValueUser = "LicenseUser";
    }
}
