using PrismTaskPanes.Controls;
using System;
using System.Linq;
using System.Runtime.InteropServices;

namespace PrismTaskPanes.Core.Extensions
{
    internal static class ControlsExtensions
    {
        #region Public Methods

        public static Uri GetHostUri()
        {
            var view = typeof(PrismTaskPanesView).Name;

            var result = new Uri(
                uriString: view,
                uriKind: UriKind.Relative);

            return result;
        }

        public static string GetProgId()
        {
            var result = default(string);

            var attributes = typeof(PrismTaskPanesHost).GetCustomAttributes(
                attributeType: typeof(ProgIdAttribute),
                inherit: true);

            if (attributes?.Any() ?? false)
            {
                result = (attributes.First() as ProgIdAttribute).Value;
            }

            return result;
        }

        #endregion Public Methods
    }
}