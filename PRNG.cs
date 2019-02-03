using System;
using System.Text;
using System.Security.Cryptography;
using System.Runtime.InteropServices;

namespace PRNG
{

    [ProgId("ClassicASP.PRNG")]

    [ClassInterface(ClassInterfaceType.AutoDual)]

    [Guid("8196270A-BA52-4B3A-96C0-2F783B2075B9")]

    [ComVisible(true)]

    public class PRNG
    {
        [ComVisible(true)]

        private const string rndChars = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ1234567890";
        RNGCryptoServiceProvider CryptoRNG = new RNGCryptoServiceProvider();

        public string RandomString(int strLen, string chars = rndChars)
        {
            StringBuilder res = new StringBuilder();
            using (CryptoRNG)
            {
                byte[] uintBuffer = new byte[sizeof(uint)];

                while (strLen-- > 0)
                {
                    CryptoRNG.GetBytes(uintBuffer);
                    uint num = BitConverter.ToUInt32(uintBuffer, 0);
                    res.Append(chars[(int)(num % (uint)chars.Length)]);
                }
            }
            return res.ToString();
        }

        public int RandomNumber(int min, int max)
        {
            max = max + 1;

            // Generate four random bytes
            byte[] four_bytes = new byte[4];
            CryptoRNG.GetBytes(four_bytes);

            // Convert the bytes to a UInt32
            UInt32 scale = BitConverter.ToUInt32(four_bytes, 0);

            // And use that to pick a random number >= min and < max
            return (int)(min + (max - min) * (scale / (uint.MaxValue + 1.0)));
        }

        public string RandomBytes(int byteLength, string returnType = "")
        {
            byte[] randomBytes = new byte[byteLength];
            CryptoRNG.GetBytes(randomBytes);

            if (returnType.Contains("hex")) return BitConverter.ToString(randomBytes).Replace("-", "");
            return Convert.ToBase64String(randomBytes);
        }

    }
}
