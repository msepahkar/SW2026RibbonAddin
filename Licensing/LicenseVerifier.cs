// File: Licensing/LicenseVerifier.cs
using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Web.Script.Serialization;

namespace Licensing
{
    public sealed class VerifyResult
    {
        public bool Success { get; private set; }
        public string Error { get; private set; }
        public VerifiedLicense License { get; private set; }

        private VerifyResult() { }

        public static VerifyResult Ok(VerifiedLicense lic)
        {
            return new VerifyResult { Success = true, License = lic, Error = null };
        }
        public static VerifyResult Fail(string error)
        {
            return new VerifyResult { Success = false, Error = error, License = null };
        }
    }

    public sealed class VerifiedLicense
    {
        public string Kid { get; internal set; }
        public string Product { get; internal set; }
        public string MachineId { get; internal set; }
        public DateTimeOffset? NotBefore { get; internal set; }
        public DateTimeOffset? Expires { get; internal set; }
        public DateTimeOffset? IssuedAt { get; internal set; }
        public IDictionary<string, object> Claims { get; internal set; }
    }

    public sealed class LicenseTrustStore
    {
        private readonly Dictionary<string, ECDsa> _byKid = new Dictionary<string, ECDsa>(StringComparer.Ordinal);
        private ECDsa _singleKeyNoKid;

        // Local Base64url decoder (so we don't depend on helpers in other classes)
        private static byte[] B64Url(string input)
        {
            if (string.IsNullOrEmpty(input)) return new byte[0];
            string s = input.Replace('-', '+').Replace('_', '/');
            switch (s.Length % 4)
            {
                case 2: s += "=="; break;
                case 3: s += "="; break;
            }
            return Convert.FromBase64String(s);
        }

        public static LicenseTrustStore FromJson(string jwkOrJwksJson)
        {
            if (string.IsNullOrWhiteSpace(jwkOrJwksJson))
                throw new ArgumentException("Empty JWK/JWKS JSON.");

            var ser = new JavaScriptSerializer();
            jwkOrJwksJson = jwkOrJwksJson.Trim();
            var store = new LicenseTrustStore();

            if (jwkOrJwksJson.Contains("\"keys\""))
            {
                var root = ser.DeserializeObject(jwkOrJwksJson) as Dictionary<string, object>;
                if (root == null || !root.ContainsKey("keys"))
                    throw new ArgumentException("Invalid JWKS JSON: missing 'keys'.");
                var arr = root["keys"] as object[];
                if (arr == null || arr.Length == 0) throw new ArgumentException("JWKS has no keys.");
                foreach (var o in arr)
                {
                    var jwk = o as Dictionary<string, object>;
                    if (jwk == null) continue;
                    store.AddFromJwk(jwk);
                }
            }
            else
            {
                var jwk = ser.DeserializeObject(jwkOrJwksJson) as Dictionary<string, object>;
                if (jwk == null) throw new ArgumentException("Invalid JWK JSON.");
                store.AddFromJwk(jwk);
            }
            return store;
        }

        private void AddFromJwk(Dictionary<string, object> jwk)
        {
            var kty = GetString(jwk, "kty");
            var crv = GetString(jwk, "crv");
            if (!string.Equals(kty, "EC", StringComparison.Ordinal) ||
                !string.Equals(crv, "P-256", StringComparison.Ordinal))
                throw new NotSupportedException("Only EC P-256 keys are supported.");

            var xB64 = GetString(jwk, "x");
            var yB64 = GetString(jwk, "y");
            if (string.IsNullOrEmpty(xB64) || string.IsNullOrEmpty(yB64))
                throw new ArgumentException("JWK must include x and y.");

            var x = B64Url(xB64);
            var y = B64Url(yB64);
            if (x.Length != 32 || y.Length != 32)
                throw new ArgumentException("JWK x/y must be 32 bytes for P-256.");

            var parameters = new ECParameters
            {
                Curve = ECCurve.CreateFromFriendlyName("nistP256"),
                Q = new ECPoint { X = x, Y = y }
            };
            var ecdsa = ECDsa.Create();
            ecdsa.ImportParameters(parameters);

            var kid = GetString(jwk, "kid");
            if (!string.IsNullOrEmpty(kid))
                _byKid[kid] = ecdsa;
            else
                _singleKeyNoKid = ecdsa;
        }

        private static string GetString(Dictionary<string, object> dict, string name)
        {
            object v;
            if (dict.TryGetValue(name, out v) && v != null)
                return Convert.ToString(v);
            return null;
        }

        internal bool TryGet(string kid, out ECDsa key)
        {
            if (!string.IsNullOrEmpty(kid) && _byKid.TryGetValue(kid, out key))
                return true;
            if (string.IsNullOrEmpty(kid) && _byKid.Count == 1 && _singleKeyNoKid == null)
            {
                key = _byKid.Values.First();
                return true;
            }
            if (_singleKeyNoKid != null && string.IsNullOrEmpty(kid))
            {
                key = _singleKeyNoKid;
                return true;
            }
            key = null;
            return false;
        }
    }

    public sealed class Es256LicenseVerifier
    {
        private readonly LicenseTrustStore _trust;
        private readonly string _expectedProduct;
        private readonly Func<string> _machineIdProvider;
        private readonly TimeSpan _clockSkew;

        public Es256LicenseVerifier(LicenseTrustStore trust, string expectedProduct,
            Func<string> machineIdProvider, TimeSpan clockSkew)
        {
            _trust = trust ?? throw new ArgumentNullException("trust");
            _expectedProduct = expectedProduct ?? "";
            _machineIdProvider = machineIdProvider ?? (() => null);
            _clockSkew = clockSkew;
        }

        public VerifyResult Verify(string compactJws)
        {
            if (string.IsNullOrWhiteSpace(compactJws))
                return VerifyResult.Fail("Empty token.");

            var parts = compactJws.Split('.');
            if (parts.Length != 3) return VerifyResult.Fail("Token is not in 'header.payload.signature' form.");

            var headerUtf8 = Base64UrlDecode(parts[0]);
            var payloadUtf8 = Base64UrlDecode(parts[1]);
            var sigRaw = Base64UrlDecode(parts[2]);

            // Parse header/payload JSON
            var ser = new JavaScriptSerializer();
            Dictionary<string, object> header, payload;
            try
            {
                header = ser.DeserializeObject(Encoding.UTF8.GetString(headerUtf8)) as Dictionary<string, object>;
                payload = ser.DeserializeObject(Encoding.UTF8.GetString(payloadUtf8)) as Dictionary<string, object>;
            }
            catch (Exception ex)
            {
                return VerifyResult.Fail("Invalid JSON: " + ex.Message);
            }
            if (header == null || payload == null)
                return VerifyResult.Fail("Invalid token JSON.");

            var alg = GetString(header, "alg");
            if (!string.Equals(alg, "ES256", StringComparison.Ordinal))
                return VerifyResult.Fail("Unsupported alg: " + alg);

            var kid = GetString(header, "kid");

            // Get key
            ECDsa key;
            if (!_trust.TryGet(kid, out key) || key == null)
                return VerifyResult.Fail(string.IsNullOrEmpty(kid) ?
                    "No matching public key." : ("Unknown kid: " + kid));

            // Verify signature
            var signingInput = Encoding.ASCII.GetBytes(parts[0] + "." + parts[1]);
            byte[] der = P1363ToDer(sigRaw); // .NET VerifyData expects DER
            bool ok;
            try
            {
                ok = key.VerifyData(signingInput, der, HashAlgorithmName.SHA256);
            }
            catch (Exception ex)
            {
                return VerifyResult.Fail("Signature check failed: " + ex.Message);
            }
            if (!ok) return VerifyResult.Fail("Invalid signature.");

            // Claims
            var lic = new VerifiedLicense();
            lic.Kid = kid;
            lic.Product = GetString(payload, "prd");
            lic.MachineId = GetString(payload, "mid");
            lic.IssuedAt = ReadUnixSeconds(payload, "iat");
            lic.NotBefore = ReadUnixSeconds(payload, "nbf");
            lic.Expires = ReadUnixSeconds(payload, "exp");
            lic.Claims = payload;

            if (!string.IsNullOrEmpty(_expectedProduct))
            {
                if (!string.Equals(lic.Product, _expectedProduct, StringComparison.Ordinal))
                    return VerifyResult.Fail("Wrong product: " + lic.Product);
            }

            // Time checks
            var now = DateTimeOffset.UtcNow;
            if (lic.NotBefore.HasValue && now + _clockSkew < lic.NotBefore.Value)
                return VerifyResult.Fail("License not yet valid.");
            if (lic.Expires.HasValue && now - _clockSkew > lic.Expires.Value)
                return VerifyResult.Fail("License expired.");

            // Machine binding
            if (!string.IsNullOrEmpty(lic.MachineId))
            {
                var current = _machineIdProvider != null ? _machineIdProvider() : null;
                if (!string.Equals(lic.MachineId, current, StringComparison.OrdinalIgnoreCase))
                    return VerifyResult.Fail("This license is bound to a different machine.");
            }

            return VerifyResult.Ok(lic);
        }

        private static string GetString(Dictionary<string, object> dict, string name)
        {
            object v;
            if (dict != null && dict.TryGetValue(name, out v) && v != null)
                return Convert.ToString(v);
            return null;
        }

        private static DateTimeOffset? ReadUnixSeconds(Dictionary<string, object> dict, string name)
        {
            object v;
            if (dict.TryGetValue(name, out v) && v != null)
            {
                try
                {
                    long s = Convert.ToInt64(v);
                    return DateTimeOffset.FromUnixTimeSeconds(s);
                }
                catch { }
            }
            return null;
        }

        // Convert IEEE P1363 (r||s) 64 bytes to ASN.1 DER sequence of two INTEGERs
        private static byte[] P1363ToDer(byte[] sig)
        {
            if (sig == null || sig.Length != 64) throw new ArgumentException("ES256 signature must be 64 bytes.");
            byte[] r = new byte[32];
            byte[] s = new byte[32];
            Buffer.BlockCopy(sig, 0, r, 0, 32);
            Buffer.BlockCopy(sig, 32, s, 0, 32);

            var derR = EncodeDerInteger(r);
            var derS = EncodeDerInteger(s);
            var lenBytes = EncodeDerLength(derR.Length + derS.Length);

            var result = new byte[1 + lenBytes.Length + derR.Length + derS.Length];
            int offset = 0;
            result[offset++] = 0x30; // SEQUENCE
            Buffer.BlockCopy(lenBytes, 0, result, offset, lenBytes.Length); offset += lenBytes.Length;
            Buffer.BlockCopy(derR, 0, result, offset, derR.Length); offset += derR.Length;
            Buffer.BlockCopy(derS, 0, result, offset, derS.Length); offset += derS.Length;
            return result;
        }

        private static byte[] EncodeDerInteger(byte[] value)
        {
            // Trim leading zeros
            int idx = 0;
            while (idx < value.Length && value[idx] == 0) idx++;
            byte[] v = (idx == value.Length) ? new byte[] { 0 } : Sub(value, idx, value.Length - idx);

            // If high bit set, prepend 0x00 to signal positive integer
            if ((v[0] & 0x80) != 0)
            {
                byte[] tmp = new byte[v.Length + 1];
                tmp[0] = 0x00;
                Buffer.BlockCopy(v, 0, tmp, 1, v.Length);
                v = tmp;
            }

            byte[] len = EncodeDerLength(v.Length);
            byte[] res = new byte[1 + len.Length + v.Length];
            int off = 0;
            res[off++] = 0x02; // INTEGER
            Buffer.BlockCopy(len, 0, res, off, len.Length); off += len.Length;
            Buffer.BlockCopy(v, 0, res, off, v.Length);
            return res;
        }

        private static byte[] EncodeDerLength(int length)
        {
            if (length < 128) return new byte[] { (byte)length };
            // Long form
            var bytes = IntToBigEndian(length);
            byte[] res = new byte[1 + bytes.Length];
            res[0] = (byte)(0x80 | bytes.Length);
            Buffer.BlockCopy(bytes, 0, res, 1, bytes.Length);
            return res;
        }

        private static byte[] IntToBigEndian(int value)
        {
            byte[] b = BitConverter.GetBytes(value);
            if (BitConverter.IsLittleEndian) Array.Reverse(b);
            // Trim leading zeros
            int i = 0; while (i < b.Length && b[i] == 0) i++;
            if (i == 0) return b;
            byte[] r = new byte[b.Length - i];
            Buffer.BlockCopy(b, i, r, 0, r.Length);
            return r;
        }

        private static byte[] Sub(byte[] src, int offset, int count)
        {
            var r = new byte[count];
            Buffer.BlockCopy(src, offset, r, 0, count);
            return r;
        }

        internal static byte[] Base64UrlDecode(string input)
        {
            if (string.IsNullOrEmpty(input)) return new byte[0];
            string s = input.Replace('-', '+').Replace('_', '/');
            switch (s.Length % 4)
            {
                case 2: s += "=="; break;
                case 3: s += "="; break;
            }
            return Convert.FromBase64String(s);
        }
    }
}
