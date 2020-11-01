#pragma warning disable CA1031 // Keine allgemeinen Ausnahmetypen abfangen

using NetOffice.OfficeApi;
using NetOffice.OfficeApi.Enums;
using Prism.Regions;
using PrismTaskPanes.Controls;
using PrismTaskPanes.Extensions;
using PrismTaskPanes.Settings;
using System;
using System.Globalization;
using System.Linq;

namespace PrismTaskPanes.Factories
{
    internal class TaskPanesFactory
    {
        #region Private Fields

        private readonly ICTPFactory ctpFactory;
        private readonly IRegionManager hostRegionManager;
        private readonly int key;
        private readonly object taskPaneWindow;

        #endregion Private Fields

        #region Public Constructors

        public TaskPanesFactory(int key, ICTPFactory ctpFactory, IRegionManager hostRegionManager, object taskPaneWindow)
        {
            this.key = key;
            this.ctpFactory = ctpFactory;
            this.hostRegionManager = hostRegionManager;
            this.taskPaneWindow = taskPaneWindow;
        }

        #endregion Public Constructors

        #region Public Methods

        public CustomTaskPane Get(TaskPaneSettings settings)
        {
            var result = GetTaskPane(settings);

            SetRegions(
                taskPane: result,
                settings: settings);

            return result;
        }

        #endregion Public Methods

        #region Private Methods

        private static Uri GetUriView()
        {
            var view = typeof(PrismTaskPanesView).Name;

            var result = new Uri(
                uriString: view,
                uriKind: UriKind.Relative);

            return result;
        }

        private CustomTaskPane GetTaskPane(TaskPaneSettings settings)
        {
            var result = ctpFactory.CreateCTP(
                cTPAxID: typeof(PrismTaskPanesHost).GetAxID(),
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

            result.VisibleStateChangeEvent += (taskPane) => DryIocProvider.OnTaskPaneChanged(taskPane);

            return result;
        }

        private Uri GetUriNavigation(TaskPaneSettings settings)
        {
            var parameter = new NavigationParameters();

            parameter.Add(
                key: DryIocProvider.WindowKey,
                value: key);

            if (!string.IsNullOrWhiteSpace(settings.NavigationValue))
            {
                parameter.Add(
                    key: DryIocProvider.NavigationKey,
                    value: settings.NavigationValue);
            }

            var uriString = settings.View.Name + parameter;

            var result = new Uri(
                uriString: uriString,
                uriKind: UriKind.Relative);

            return result;
        }

        private void SetRegions(CustomTaskPane taskPane, TaskPaneSettings settings)
        {
            var host = taskPane.ContentControl as PrismTaskPanesHost;
            var hostRegionName = host.GetHashCode()
                .ToString(CultureInfo.InvariantCulture);

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

            var source = GetUriNavigation(settings);

            var viewRegionManager = view.LocalRegionManager;
            viewRegionManager.RequestNavigate(
                regionName: settings.RegionName,
                source: source);
        }

        #endregion Private Methods
    }
}

#pragma warning restore CA1031 // Keine allgemeinen Ausnahmetypen abfangen