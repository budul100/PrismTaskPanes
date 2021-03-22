using Prism.Regions;
using PrismTaskPanes.Controls;
using PrismTaskPanes.Settings;
using System;

namespace PrismTaskPanes.Extensions
{
    internal static class TaskPaneExtensions
    {
        #region Public Methods

        public static Uri GetUriHost()
        {
            var view = typeof(PrismTaskPanesView).Name;

            var result = new Uri(
                uriString: view,
                uriKind: UriKind.Relative);

            return result;
        }

        public static Uri GetUriView(this TaskPaneSettings settings, int windowKey)
        {
            var parameter = new NavigationParameters();

            parameter.Add(
                key: BaseProvider.WindowKey,
                value: windowKey);

            if (!string.IsNullOrWhiteSpace(settings.NavigationValue))
            {
                parameter.Add(
                    key: BaseProvider.NavigationKey,
                    value: settings.NavigationValue);
            }

            var uriString = settings.View.Name + parameter;

            var result = new Uri(
                uriString: uriString,
                uriKind: UriKind.Relative);

            return result;
        }

        #endregion Public Methods
    }
}