using NetOffice.OfficeApi;
using NetOffice.OfficeApi.Enums;
using Prism.Regions;
using PrismTaskPanes.Controls;
using PrismTaskPanes.Configurations;
using System;
using System.Linq;
using System.Runtime.InteropServices;

namespace PrismTaskPanes.TaskPanes
{
    internal class TaskPanesFactory
    {
        #region Private Fields

        private readonly ICTPFactory ctpFactory;
        private readonly IRegionManager hostRegionManager;
        private readonly object taskPaneWindow;

        #endregion Private Fields

        #region Public Constructors

        public TaskPanesFactory(ICTPFactory ctpFactory, IRegionManager hostRegionManager, object taskPaneWindow)
        {
            this.ctpFactory = ctpFactory;
            this.hostRegionManager = hostRegionManager;
            this.taskPaneWindow = taskPaneWindow;
        }

        #endregion Public Constructors

        #region Public Methods

        public CustomTaskPane Get(Configuration settings)
        {
            var result = GetTaskPane(settings);

            SetRegions(
                taskPane: result,
                settings: settings);

            return result;
        }

        #endregion Public Methods

        #region Private Methods

        private static string GetAxID<T>()
        {
            var attributes = typeof(T).GetCustomAttributes(
                attributeType: typeof(ProgIdAttribute),
                inherit: false) as ProgIdAttribute[];

            return attributes?.Any() ?? false
                ? attributes.First().Value
                : string.Empty;
        }

        private static Uri GetUriNavigation(Configuration settings)
        {
            var view = settings.View.Name;

            var parameter = new NavigationParameters();

            if (!string.IsNullOrWhiteSpace(settings.NavigationValue))
            {
                parameter.Add(
                    key: settings.NavigationKey,
                    value: settings.NavigationValue);
            }

            var result = new Uri(
                uriString: view + parameter,
                uriKind: UriKind.Relative);

            return result;
        }

        private static Uri GetUriView()
        {
            var view = typeof(PrismTaskPanesView).Name;

            var result = new Uri(
                uriString: view,
                uriKind: UriKind.Relative);

            return result;
        }

        private CustomTaskPane GetTaskPane(Configuration settings)
        {
            var result = ctpFactory.CreateCTP(
                cTPAxID: GetAxID<PrismTaskPanesHost>(),
                cTPTitle: settings.Title,
                cTPParentWindow: taskPaneWindow) as CustomTaskPane;

            try
            {
                result.Visible = false;
                result.DockPosition = settings.DockPosition;
                result.DockPositionRestrict = settings.DockRestriction;

                if (result.DockPosition != MsoCTPDockPosition.msoCTPDockPositionLeft &&
                    result.DockPosition != MsoCTPDockPosition.msoCTPDockPositionRight)
                    result.Height = settings.Height;

                if (result.DockPosition != MsoCTPDockPosition.msoCTPDockPositionBottom &&
                    result.DockPosition != MsoCTPDockPosition.msoCTPDockPositionTop)
                    result.Width = settings.Width;
            }
            catch
            { }

            result.VisibleStateChangeEvent += (tp) => TaskPanesProvider
                .OnTaskPaneChangedEvent?
                .Invoke(
                    sender: tp,
                    e: null);

            return result;
        }

        private void SetRegions(CustomTaskPane taskPane, Configuration settings)
        {
            var host = taskPane.ContentControl as PrismTaskPanesHost;
            var hostRegionName = host.GetHashCode().ToString();

            host.SetLocalRegion(
                regionName: hostRegionName,
                regionManager: hostRegionManager);

            var uriView = GetUriView();

            hostRegionManager.RequestNavigate(
                regionName: hostRegionName,
                source: uriView);

            var view = hostRegionManager.Regions[hostRegionName].Views
                .SingleOrDefault() as PrismTaskPanesView;

            view.SetLocalRegion(
                regionName: settings.RegionName,
                regionContext: settings.RegionContext);

            var viewRegionManager = view.LocalRegionManager;
            viewRegionManager.RequestNavigate(
                regionName: settings.RegionName,
                source: GetUriNavigation(settings));
        }

        #endregion Private Methods
    }
}