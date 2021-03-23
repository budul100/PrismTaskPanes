using NetOffice.OfficeApi.Enums;
using Prism.Regions;
using PrismTaskPanes.Attributes;
using PrismTaskPanes.Commons.Enums;
using PrismTaskPanes.Host;
using PrismTaskPanes.Settings;
using System;

namespace PrismTaskPanes.Extensions
{
    internal static class TaskPaneExtensions
    {
        #region Public Methods

        public static MsoCTPDockPosition GetDockPosition(this PrismTaskPaneAttribute attribute)
        {
            switch (attribute.DockPosition)
            {
                case DockPosition.Right:
                    return MsoCTPDockPosition.msoCTPDockPositionRight;

                case DockPosition.Left:
                    return MsoCTPDockPosition.msoCTPDockPositionLeft;

                case DockPosition.Top:
                    return MsoCTPDockPosition.msoCTPDockPositionTop;

                case DockPosition.Bottom:
                    return MsoCTPDockPosition.msoCTPDockPositionBottom;

                default:
                    return MsoCTPDockPosition.msoCTPDockPositionFloating;
            }
        }

        public static MsoCTPDockPositionRestrict GetDockRestriction(this PrismTaskPaneAttribute attribute)
        {
            switch (attribute.DockRestriction)
            {
                case DockRestriction.None:
                    return MsoCTPDockPositionRestrict.msoCTPDockPositionRestrictNone;

                case DockRestriction.NoHorizontal:
                    return MsoCTPDockPositionRestrict.msoCTPDockPositionRestrictNoHorizontal;

                case DockRestriction.NoVertical:
                    return MsoCTPDockPositionRestrict.msoCTPDockPositionRestrictNoVertical;

                default:
                    return MsoCTPDockPositionRestrict.msoCTPDockPositionRestrictNoChange;
            }
        }

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