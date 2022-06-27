using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;

namespace PrismTaskPanes.Extensions
{
    [ComVisible(false)]
    public static class HashExtensions
    {
        #region Private Fields

        private const string AllCharacters = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ";
        private const int HashLength = 20;

        #endregion Private Fields

        #region Public Methods

        public static string GetHashString(this string value, int length = HashLength)
        {
            var result = value.GetHashString(
                length: length,
                chars: AllCharacters);

            return result;
        }

        #endregion Public Methods

        #region Private Methods

        private static string GetHashString(this string value, int length, string chars)
        {
            var result = default(string);

            if (!string.IsNullOrWhiteSpace(value))
            {
                using var hashString = SHA512.Create();

                var bytes = Encoding.UTF8.GetBytes(value);

                var hash1 = hashString.ComputeHash(bytes);
                var hash2 = new char[length];

                for (var i = 0; i < hash2.Length; i++)
                {
                    hash2[i] = chars[hash1[i] % chars.Length];
                }

                result = new string(hash2);
            }

            return result;
        }

        #endregion Private Methods
    }
}