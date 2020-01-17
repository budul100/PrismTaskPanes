using PrismTaskPanes.Interfaces;
using System.Text;

namespace PrismTaskPanes.Extensions
{
    internal static class HashExtensions
    {
        #region Public Methods

        public static int GetReceiverHash(this ITaskPanesReceiver receiver, string id)
        {
            var result = GetStaticHash(
                id,
                receiver.GetType().FullName);

            return result;
        }

        public static int GetStaticHash(this object value)
        {
            var result = 0;
            var text = value?.ToString();

            if (!string.IsNullOrWhiteSpace(text))
            {
                uint hash = 0;

                // if you care this can be done much faster with unsafe
                // using fixed char* reinterpreted as a byte*
                foreach (var @byte in Encoding.Unicode.GetBytes(text))
                {
                    hash += @byte;
                    hash += (hash << 10);
                    hash ^= (hash >> 6);
                }

                // final avalanche
                hash += (hash << 3);
                hash ^= (hash >> 11);
                hash += (hash << 15);

                // helpfully we only want positive integer < MUST_BE_LESS_THAN
                // so simple truncate cast is ok if not perfect
                result = (int)hash;
            }

            return result;
        }

        #endregion Public Methods

        #region Private Methods

        private static int GetStaticHash(params object[] values)
        {
            unchecked
            {
                var result = 17;

                foreach (var value in values)
                {
                    result = result * 31 + value.GetStaticHash();
                }

                return result;
            }
        }

        #endregion Private Methods
    }
}