using System;
using System.Linq;
using System.Runtime.InteropServices;

namespace PrismTaskPanes.Extensions
{
    internal static class TaskPaneExtensions
    {
        #region Public Methods

        public static string GetAxID(this Type type)
        {
            var attributes = type.GetCustomAttributes(
                attributeType: typeof(ProgIdAttribute),
                inherit: false) as ProgIdAttribute[];

            var result = attributes?.Any() ?? false
                ? attributes.First().Value
                : string.Empty;

            return result;
        }

        #endregion Public Methods
    }
}