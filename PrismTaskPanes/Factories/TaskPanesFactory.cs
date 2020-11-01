#pragma warning disable CA1031 // Keine allgemeinen Ausnahmetypen abfangen

using NetOffice.OfficeApi;
using NetOffice.OfficeApi.Enums;
using Prism.Regions;
using PrismTaskPanes.Controls;
using PrismTaskPanes.Exceptions;
using PrismTaskPanes.Extensions;
using PrismTaskPanes.Settings;
using System;
using System.Globalization;
using System.Linq;
using System.Threading.Tasks;

namespace PrismTaskPanes.Factories
{
    internal class TaskPanesFactory
    {
        #region Private Fields

        private readonly ICTPFactory ctpFactory;
        private readonly IRegionManager hostRegionManager;
        private readonly int windowKey;
        private readonly object taskPaneWindow;

        #endregion Private Fields

        #region Public Constructors

        public TaskPanesFactory(int key, ICTPFactory ctpFactory, IRegionManager hostRegionManager, object taskPaneWindow)
        {
            if (Type.GetTypeFromProgID(PrismTaskPanesHost.ProgId) == default)
            {
                throw new LibraryNotRegisteredException();
            }

            this.windowKey = key;
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



        private void SetRegions(CustomTaskPane taskPane, TaskPaneSettings settings)
        {
            var host = taskPane.ContentControl as PrismTaskPanesHost;
            var hostRegionName = host.GetHashCode()
                .ToString(CultureInfo.InvariantCulture);

            host.SetLocalRegion(
                regionName: hostRegionName,
                regionManager: hostRegionManager);

            var horstUri = TaskPaneExtensions.GetUriHost();

            hostRegionManager.RequestNavigate(
                regionName: hostRegionName,
                source: horstUri);

            var view = hostRegionManager.Regions[hostRegionName].Views
                .SingleOrDefault() as PrismTaskPanesView;

            view.SetLocalRegion(
                regionName: settings.RegionName,
                regionContext: settings.RegionContext);

            var viewUri = settings.GetUriView(windowKey);
            var viewRegionManager = view.LocalRegionManager;

            viewRegionManager.RequestNavigate(
                regionName: settings.RegionName,
                source: viewUri);
        }

        #endregion Private Methods
    }
}

#pragma warning restore CA1031 // Keine allgemeinen Ausnahmetypen abfangen