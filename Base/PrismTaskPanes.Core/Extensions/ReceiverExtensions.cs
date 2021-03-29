using PrismTaskPanes.Attributes;
using PrismTaskPanes.Extensions;
using PrismTaskPanes.Interfaces;
using System.Collections.Generic;
using System.Runtime.InteropServices;

namespace PrismTaskPanes.Core.Extensions
{
    [ComVisible(false)]
    public static class ReceiverExtensions
    {
        #region Private Fields

        private const string Separator = "\n";

        #endregion Private Fields

        #region Public Methods

        public static IEnumerable<PrismTaskPaneAttribute> GetAttributes(this ITaskPanesReceiver receiver)
        {
            var type = receiver.GetType();

            var attributes = type.GetCustomAttributes(
                attributeType: typeof(PrismTaskPaneAttribute),
                inherit: true);

            foreach (var attribute in attributes)
            {
                var result = attribute as PrismTaskPaneAttribute;

                result.ReceiverHash = receiver.GetReceiverHash(result.ID);

                yield return result;
            }
        }

        public static string GetReceiverHash(this ITaskPanesReceiver receiver, string id)
        {
            var value = string.Join(
                Separator,
                id,
                receiver.GetType().FullName);

            var result = value.GetHashString();

            return result;
        }

        #endregion Public Methods
    }
}