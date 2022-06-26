using NetOffice.OfficeApi;
using NetOffice.OfficeApi.Enums;
using Prism.Regions;
using PrismTaskPanes.Constants;
using PrismTaskPanes.Enums;
using PrismTaskPanes.Settings;
using System;
using System.Runtime.InteropServices;

namespace PrismTaskPanes.Extensions
{
    [ComVisible(false)]
    internal static class TaskPaneExtensions
    {
        #region Public Methods

        public static DockPosition GetDockPosition(this CustomTaskPane taskPane)
        {
            return taskPane.DockPosition switch
            {
                MsoCTPDockPosition.msoCTPDockPositionRight => DockPosition.Right,
                MsoCTPDockPosition.msoCTPDockPositionLeft => DockPosition.Left,
                MsoCTPDockPosition.msoCTPDockPositionTop => DockPosition.Top,
                MsoCTPDockPosition.msoCTPDockPositionBottom => DockPosition.Bottom,
                _ => DockPosition.Floating,
            };
        }

        public static MsoCTPDockPosition GetDockPosition(this TaskPaneSettings settings)
        {
            return settings.DockPosition switch
            {
                DockPosition.Right => MsoCTPDockPosition.msoCTPDockPositionRight,
                DockPosition.Left => MsoCTPDockPosition.msoCTPDockPositionLeft,
                DockPosition.Top => MsoCTPDockPosition.msoCTPDockPositionTop,
                DockPosition.Bottom => MsoCTPDockPosition.msoCTPDockPositionBottom,
                _ => MsoCTPDockPosition.msoCTPDockPositionFloating,
            };
        }

        public static MsoCTPDockPositionRestrict GetDockRestriction(this TaskPaneSettings settings)
        {
            return settings.DockRestriction switch
            {
                DockRestriction.None => MsoCTPDockPositionRestrict.msoCTPDockPositionRestrictNone,
                DockRestriction.NoHorizontal => MsoCTPDockPositionRestrict.msoCTPDockPositionRestrictNoHorizontal,
                DockRestriction.NoVertical => MsoCTPDockPositionRestrict.msoCTPDockPositionRestrictNoVertical,
                _ => MsoCTPDockPositionRestrict.msoCTPDockPositionRestrictNoChange,
            };
        }

        public static Uri GetUriView(this TaskPaneSettings settings, int windowKey)
        {
            var parameter = new NavigationParameters();

            parameter.Add(
                key: ProviderConstants.WindowKey,
                value: windowKey);

            if (!string.IsNullOrWhiteSpace(settings.NavigationValue))
            {
                parameter.Add(
                    key: ProviderConstants.NavigationKey,
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