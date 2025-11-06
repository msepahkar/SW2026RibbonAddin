using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Numerics;
using System.Security.Cryptography;
using System.Text;
using System.Web.Script.Serialization;

namespace SW2025RibbonAddin.Licensing
{
    public sealed class VerifiedLicense
    {
        public string Token { get; set; }
        public string KeyId { get; set; }
        public string Product { get; set; }
        public string MachineId { get; set; }
        public string UserName { get; set; }
        public DateTimeOffset? IssuedAt { get; set; }
        public DateTimeOffset? NotBefore { get; set; }
        public DateTimeOffset? Expires { get; set; }
        public string PayloadJson { get; set; }
    }

    internal sealed class EcPublicJwk
    {
        public string Kty; public string Crv; public string X; public string Y; public string Kid;
    }

    public sealed class LicenseVerifier
    {
        private readonly Dictionary<string, ECDsa> _keysByKid;
        private readonly ECDsa _singleKey;
        private readonly byte[] _qx;  // for pure verifier (single key case)
        private readonly byte[] _qy;

        private readonly string _expectedProduct;
        private readonly string _expectedMachineId;
        private readonly TimeSpan _skew;
        private readonly TimeSpan _grace;

        // secp256r1 order (n) and n/2 (big‑endian) for low‑S normalization
        private static readonly byte[] N_be = FromHex("FFFFFFFF00000000FFFFFFFFFFFFFFFFBCE6FAADA7179E84F3B9CAC2FC632551");
        private static readonly byte[] NHalf_be = FromHex("7FFFFFFF800000007FFFFFFFFFFFFFFFDE737D56D38BCF4279DCE5617E3192A8");

        public LicenseVerifier(string jwkOrJwksJson,
                               string expectedProduct,
                               string expectedMachineId,
                               TimeSpan allowedClockSkew,
                               TimeSpan expiryGrace)
        {
            _expectedProduct = expectedProduct ?? "";
            _expectedMachineId = expectedMachineId ?? "";
            _skew = allowedClockSkew;
            _grace = expiryGrace;

            var keys = ParseJwks(jwkOrJwksJson);
            _keysByKid = new Dictionary<string, ECDsa>(StringComparer.Ordinal);

            foreach (var jwk in keys)
            {
                var x = B64UrlDecode(jwk.X);
                var y = B64UrlDecode(jwk.Y);

                var ec = ECDsa.Create();
                ec.ImportParameters(new ECParameters
                {
                    Curve = ECCurve.NamedCurves.nistP256,
                    Q = new ECPoint { X = x, Y = y }
                });

                _keysByKid[jwk.Kid ?? ""] = ec;
            }

            if (_keysByKid.Count == 1)
            {
                _singleKey = _keysByKid.Values.First();
                // cache x,y for pure fallback
                var only = keys[0];
                _qx = B64UrlDecode(only.X);
                _qy = B64UrlDecode(only.Y);
            }
        }

        public bool TryVerify(string compactJws, out VerifiedLicense license, out string error)
        {
            license = null; error = null;

            var parts = (compactJws ?? "").Split('.');
            if (parts.Length != 3) { error = "Token is not in compact JWS format."; return false; }

            byte[] headerBytes, payloadBytes, sigBytes;
            try
            {
                headerBytes = B64UrlDecode(parts[0]);
                payloadBytes = B64UrlDecode(parts[1]);
                sigBytes = B64UrlDecode(parts[2]);
            }
            catch { error = "Invalid base64url sections."; return false; }

            var js = new JavaScriptSerializer();
            Dictionary<string, object> header, payload;
            try
            {
                header = js.Deserialize<Dictionary<string, object>>(Encoding.UTF8.GetString(headerBytes));
                payload = js.Deserialize<Dictionary<string, object>>(Encoding.UTF8.GetString(payloadBytes));
            }
            catch { error = "Malformed JSON in token."; return false; }

            if (!string.Equals(GetString(header, "alg"), "ES256", StringComparison.Ordinal))
            { error = "Unsupported alg (expected ES256)."; return false; }

            var kid = GetString(header, "kid");
            ECDsa key = null;
            if (!string.IsNullOrEmpty(kid))
            {
                if (!_keysByKid.TryGetValue(kid, out key)) { error = "Unknown key id."; return false; }
            }
            else
            {
                if (_keysByKid.Count != 1) { error = "Token has no 'kid' and multiple keys are configured."; return false; }
                key = _singleKey;
            }

            // Reconstruct signing input
            var signingInput = Encoding.ASCII.GetBytes(parts[0] + "." + parts[1]);

            // Convert to DER with low‑S (compat path)
            byte[] derSig = sigBytes.Length == 64 ? P1363ToDerLowS(sigBytes) : EnsureLowSIfDer(sigBytes);

            // ---- 1) Verify using platform API (most compatible) ----
            bool ok = false;

            // older stacks: VerifyHash first
            try
            {
                using (var sha = SHA256.Create())
                {
                    var hash = sha.ComputeHash(signingInput);
                    ok = key.VerifyHash(hash, derSig);
                }
            }
            catch { ok = false; }

            // newer stacks: VerifyData
            if (!ok)
            {
                try { ok = key.VerifyData(signingInput, derSig, HashAlgorithmName.SHA256); }
                catch { ok = false; }
            }

            // ---- 2) Pure C# verifier fallback (IEEE‑P1363, original r|s) ----
            if (!ok)
            {
                // Need X,Y for pure path — available when a single key is configured or kid matched one of ours
                byte[] qx = _qx, qy = _qy;
                if (string.IsNullOrEmpty(kid) && qx == null) { /* single key cached into _qx/_qy */ }
                else if (!string.IsNullOrEmpty(kid))
                {
                    // Re-extract x,y for the matched key
                    foreach (var kv in _keysByKid)
                    {
                        if (kv.Key == kid)
                        {
                            // rebuild EC parameters to fetch Q.X/Q.Y (not exposed directly)
                            try
                            {
                                var p = kv.Value.ExportParameters(false);
                                qx = p.Q.X; qy = p.Q.Y;
                            }
                            catch { qx = _qx; qy = _qy; }
                            break;
                        }
                    }
                }

                if (qx != null && qy != null && sigBytes.Length == 64)
                    ok = VerifyP256_P1363_Pure(signingInput, sigBytes, qx, qy);
            }

            if (!ok) { error = "Invalid signature."; return false; }

            // ---- Claims checks ----
            var product = GetString(payload, "prd");
            var machineId = GetString(payload, "mid");
            var userName = GetString(payload, "usr"); // optional

            if (!string.Equals(product, _expectedProduct, StringComparison.Ordinal))
            { error = "Wrong product."; return false; }

            if (!string.Equals((machineId ?? "").Trim(), (_expectedMachineId ?? "").Trim(), StringComparison.OrdinalIgnoreCase))
            { error = "Machine Id mismatch."; return false; }

            var now = DateTimeOffset.UtcNow;

            DateTimeOffset? iat = TryGetUnix(payload, "iat");
            if (iat.HasValue && iat.Value - now > _skew)
            { error = "Token 'iat' is in the future."; return false; }

            DateTimeOffset? nbf = TryGetUnix(payload, "nbf");
            if (nbf.HasValue && now + _skew < nbf.Value)
            { error = "License not valid yet."; return false; }

            DateTimeOffset? exp = TryGetUnix(payload, "exp");
            if (exp.HasValue && now - _grace > exp.Value)
            { error = "License expired."; return false; }

            license = new VerifiedLicense
            {
                Token = compactJws,
                KeyId = !string.IsNullOrEmpty(kid) ? kid : _keysByKid.Keys.First(),
                Product = product,
                MachineId = machineId,
                UserName = userName,
                IssuedAt = iat,
                NotBefore = nbf,
                Expires = exp,
                PayloadJson = Encoding.UTF8.GetString(payloadBytes)
            };
            return true;
        }

        // ====================== Pure C# P‑256 verify (IEEE‑P1363) ======================

        private static bool VerifyP256_P1363_Pure(byte[] data, byte[] sigP1363, byte[] qx32, byte[] qy32)
        {
            if (sigP1363 == null || sigP1363.Length != 64) return false;

            // secp256r1 parameters
            BigInteger p = Big("FFFFFFFF00000001000000000000000000000000FFFFFFFFFFFFFFFFFFFFFFFF");
            BigInteger a = Big("FFFFFFFF00000001000000000000000000000000FFFFFFFFFFFFFFFFFFFFFFFC");
            BigInteger b = Big("5AC635D8AA3A93E7B3EBBD55769886BC651D06B0CC53B0F63BCE3C3E27D2604B");
            BigInteger Gx = Big("6B17D1F2E12C4247F8BCE6E563A440F277037D812DEB33A0F4A13945D898C296");
            BigInteger Gy = Big("4FE342E2FE1A7F9B8EE7EB4A7C0F9E162BCE33576B315ECECBB6406837BF51F5");
            BigInteger n = Big("FFFFFFFF00000000FFFFFFFFFFFFFFFFBCE6FAADA7179E84F3B9CAC2FC632551");

            BigInteger r = ToBig(sigP1363, 0, 32);
            BigInteger s = ToBig(sigP1363, 32, 32);
            if (r.Sign <= 0 || r >= n || s.Sign <= 0 || s >= n) return false;

            byte[] hash; using (var sha = SHA256.Create()) hash = sha.ComputeHash(data);
            BigInteger z = ToBig(hash, 0, hash.Length);

            BigInteger w = ModInverse(s, n);
            BigInteger u1 = (z * w) % n; if (u1.Sign < 0) u1 += n;
            BigInteger u2 = (r * w) % n;

            var G = new ECPointBI(Gx, Gy);
            var Q = new ECPointBI(ToBig(qx32, 0, 32), ToBig(qy32, 0, 32));

            var X = Add(ScalarMult(u1, G, a, p), ScalarMult(u2, Q, a, p), a, p);
            if (X.Inf) return false;

            BigInteger v = X.X % n; if (v.Sign < 0) v += n;
            return v == r;
        }

        private struct ECPointBI
        {
            public BigInteger X, Y; public bool Inf;
            public ECPointBI(BigInteger x, BigInteger y) { X = x; Y = y; Inf = false; }
            public static ECPointBI Infinity => new ECPointBI { Inf = true, X = BigInteger.Zero, Y = BigInteger.Zero };
        }

        private static ECPointBI Add(ECPointBI P, ECPointBI Q, BigInteger a, BigInteger p)
        {
            if (P.Inf) return Q;
            if (Q.Inf) return P;
            if (P.X == Q.X && (P.Y + Q.Y) % p == 0) return ECPointBI.Infinity;

            BigInteger lambda;
            if (!(P.X == Q.X && P.Y == Q.Y))
                lambda = ((Q.Y - P.Y) * ModInverse((Q.X - P.X) % p, p)) % p;
            else
                lambda = ((3 * P.X * P.X + a) * ModInverse((2 * P.Y) % p, p)) % p;

            if (lambda.Sign < 0) lambda += p;

            BigInteger xr = (lambda * lambda - P.X - Q.X) % p; if (xr.Sign < 0) xr += p;
            BigInteger yr = (lambda * (P.X - xr) - P.Y) % p; if (yr.Sign < 0) yr += p;
            return new ECPointBI(xr, yr);
        }

        private static ECPointBI ScalarMult(BigInteger k, ECPointBI P, BigInteger a, BigInteger p)
        {
            ECPointBI N = ECPointBI.Infinity, Q = P;
            while (k > 0)
            {
                if (!k.IsEven) N = Add(N, Q, a, p);
                Q = Add(Q, Q, a, p);
                k >>= 1;
            }
            return N;
        }

        private static BigInteger ModInverse(BigInteger a, BigInteger m)
        {
            BigInteger t = 0, newt = 1;
            BigInteger r = m, newr = a % m; if (newr.Sign < 0) newr += m;
            while (newr != 0)
            {
                BigInteger q = r / newr;
                var tmpT = newt; newt = t - q * newt; t = tmpT;
                var tmpR = newr; newr = r - q * newr; r = tmpR;
            }
            if (r != 1) throw new InvalidOperationException("mod inverse does not exist");
            if (t.Sign < 0) t += m;
            return t;
        }

        private static BigInteger Big(string hex) => BigInteger.Parse("0" + hex, System.Globalization.NumberStyles.HexNumber);

        private static BigInteger ToBig(byte[] be, int offset, int count)
        {
            var tmp = new byte[count + 1];
            for (int i = 0; i < count; i++) tmp[count - 1 - i] = be[offset + i];
            return new BigInteger(tmp);
        }

        // ====================== helpers ======================

        private static string GetString(Dictionary<string, object> obj, string name)
        {
            object v; return (obj != null && obj.TryGetValue(name, out v) && v != null) ? v.ToString() : "";
        }

        private static DateTimeOffset? TryGetUnix(Dictionary<string, object> obj, string name)
        {
            object v;
            if (obj == null || !obj.TryGetValue(name, out v) || v == null) return null;
            try { long s; if (long.TryParse(v.ToString(), out s)) return DateTimeOffset.FromUnixTimeSeconds(s); }
            catch { }
            return null;
        }

        private static List<EcPublicJwk> ParseJwks(string jsonString)
        {
            var js = new JavaScriptSerializer();
            var list = new List<EcPublicJwk>();

            var root = js.DeserializeObject(jsonString);
            var dict = root as Dictionary<string, object>;

            if (dict != null && dict.ContainsKey("keys"))
            {
                IEnumerable items = null;
                var keysObj = dict["keys"];
                if (keysObj is object[]) items = (object[])keysObj;
                else if (keysObj is ArrayList) items = (ArrayList)keysObj;

                if (items != null)
                {
                    foreach (var item in items)
                    {
                        var jwk = TryReadJwk(item as Dictionary<string, object>);
                        if (jwk != null) list.Add(jwk);
                    }
                }
            }
            else
            {
                var jwk = TryReadJwk(dict);
                if (jwk != null) list.Add(jwk);
            }

            if (list.Count == 0)
                throw new InvalidOperationException("TrustedKeysJson does not contain a valid EC JWK/JWKS.");

            return list;
        }

        private static EcPublicJwk TryReadJwk(Dictionary<string, object> d)
        {
            if (d == null) return null;
            var kty = GetString(d, "kty"); if (!string.Equals(kty, "EC", StringComparison.OrdinalIgnoreCase)) return null;
            var crv = GetString(d, "crv"); if (!string.Equals(crv, "P-256", StringComparison.OrdinalIgnoreCase)) throw new NotSupportedException("Only P-256 is supported.");
            var x = GetString(d, "x"); var y = GetString(d, "y");
            if (string.IsNullOrWhiteSpace(x) || string.IsNullOrWhiteSpace(y)) return null;

            return new EcPublicJwk { Kty = "EC", Crv = crv, X = x, Y = y, Kid = GetString(d, "kid") };
        }

        private static byte[] B64UrlDecode(string s)
        {
            if (s == null) return new byte[0];
            s = s.Replace('-', '+').Replace('_', '/');
            switch (s.Length % 4) { case 2: s += "=="; break; case 3: s += "="; break; }
            return Convert.FromBase64String(s);
        }

        // ----- low‑S normalization + DER helpers -----

        private static byte[] P1363ToDerLowS(byte[] sig64)
        {
            var r = new byte[32];
            var s = new byte[32];
            Buffer.BlockCopy(sig64, 0, r, 0, 32);
            Buffer.BlockCopy(sig64, 32, s, 0, 32);

            if (CompareBigEndian(s, NHalf_be) > 0)
                s = SubtractBigEndian(N_be, s); // s = n - s

            r = TrimZeros(r); s = TrimZeros(s);
            if (r.Length > 0 && (r[0] & 0x80) != 0) r = PrependZero(r);
            if (s.Length > 0 && (s[0] & 0x80) != 0) s = PrependZero(s);

            var len = 2 + r.Length + 2 + s.Length;
            var der = new byte[2 + len];
            int i = 0;
            der[i++] = 0x30; der[i++] = (byte)len;
            der[i++] = 0x02; der[i++] = (byte)r.Length; Buffer.BlockCopy(r, 0, der, i, r.Length); i += r.Length;
            der[i++] = 0x02; der[i++] = (byte)s.Length; Buffer.BlockCopy(s, 0, der, i, s.Length);
            return der;
        }

        private static byte[] EnsureLowSIfDer(byte[] der)
        {
            try
            {
                int i = 0;
                if (der[i++] != 0x30) return der;
                int seqLen = der[i++];

                if (der[i++] != 0x02) return der;
                int rLen = der[i++];
                var r = new byte[rLen]; Buffer.BlockCopy(der, i, r, 0, rLen); i += rLen;

                if (der[i++] != 0x02) return der;
                int sLen = der[i++];
                var s = new byte[sLen]; Buffer.BlockCopy(der, i, s, 0, sLen);

                r = TrimZeros(r); s = TrimZeros(s);
                r = PadLeft32(r); s = PadLeft32(s);

                if (CompareBigEndian(s, NHalf_be) > 0)
                    s = SubtractBigEndian(N_be, s);

                r = TrimZeros(r); s = TrimZeros(s);
                if (r.Length > 0 && (r[0] & 0x80) != 0) r = PrependZero(r);
                if (s.Length > 0 && (s[0] & 0x80) != 0) s = PrependZero(s);

                var len = 2 + r.Length + 2 + s.Length;
                var outDer = new byte[2 + len];
                int j = 0;
                outDer[j++] = 0x30; outDer[j++] = (byte)len;
                outDer[j++] = 0x02; outDer[j++] = (byte)r.Length; Buffer.BlockCopy(r, 0, outDer, j, r.Length); j += r.Length;
                outDer[j++] = 0x02; outDer[j++] = (byte)s.Length; Buffer.BlockCopy(s, 0, outDer, j, s.Length);
                return outDer;
            }
            catch { return der; }
        }

        private static byte[] TrimZeros(byte[] v)
        {
            int i = 0; while (i < v.Length - 1 && v[i] == 0) i++;
            if (i == 0) return v;
            var o = new byte[v.Length - i]; Buffer.BlockCopy(v, i, o, 0, o.Length); return o;
        }

        private static byte[] PrependZero(byte[] v)
        {
            var o = new byte[v.Length + 1]; Buffer.BlockCopy(v, 0, o, 1, v.Length); return o;
        }

        private static byte[] PadLeft32(byte[] v)
        {
            if (v.Length >= 32) return v;
            var o = new byte[32]; Buffer.BlockCopy(v, 0, o, 32 - v.Length, v.Length); return o;
        }

        private static int CompareBigEndian(byte[] a, byte[] b)
        {
            int len = Math.Max(a.Length, b.Length);
            for (int i = 0; i < len; i++)
            {
                int ai = i >= len - a.Length ? a[i - (len - a.Length)] : 0;
                int bi = i >= len - b.Length ? b[i - (len - b.Length)] : 0;
                if (ai != bi) return ai > bi ? 1 : -1;
            }
            return 0;
        }

        private static byte[] SubtractBigEndian(byte[] a, byte[] b)
        {
            int len = Math.Max(a.Length, b.Length);
            var aa = new byte[len]; Buffer.BlockCopy(a, 0, aa, len - a.Length, a.Length);
            var bb = new byte[len]; Buffer.BlockCopy(b, 0, bb, len - b.Length, b.Length);

            var r = new byte[len]; int borrow = 0;
            for (int i = len - 1; i >= 0; i--)
            {
                int v = aa[i] - bb[i] - borrow;
                if (v < 0) { v += 256; borrow = 1; } else borrow = 0;
                r[i] = (byte)v;
            }
            return TrimZeros(r);
        }

        private static byte[] FromHex(string hex)
        {
            int len = hex.Length / 2;
            var r = new byte[len];
            for (int i = 0; i < len; i++)
                r[i] = Convert.ToByte(hex.Substring(2 * i, 2), 16);
            return r;
        }
    }
}
