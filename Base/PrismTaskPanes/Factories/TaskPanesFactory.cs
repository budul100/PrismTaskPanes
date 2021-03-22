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
using System.Runtime.InteropServices;
using System.Windows.Controls;
using System.Windows.Forms;
using System.Windows.Forms.Integration;

namespace PrismTaskPanes.Factories
{
    [ComVisible(false)]
    public class TaskPanesFactory
    {
        #region Private Fields

        private const string ProgID = "PrismTaskPanes.Controls.PrismTaskPanesHost";

        private readonly ICTPFactory ctpFactory;
        private readonly IRegionManager hostRegionManager;
        private readonly object taskPaneWindow;
        private readonly int windowKey;

        #endregion Private Fields

        #region Public Constructors

        public TaskPanesFactory(int key, ICTPFactory ctpFactory, IRegionManager hostRegionManager, object taskPaneWindow)
        {
            if (Type.GetTypeFromProgID(ProgID) == default)
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

            var hostRegion = GetHostRegion(result);
            var hostView = GetHostView(hostRegion);

            SetRegions(
                hostView: hostView,
                settings: settings);

            return result;
        }

        #endregion Public Methods

        #region Private Methods

        private string GetHostRegion(CustomTaskPane taskPane)
        {
            var host = new ContentControl();

            var result = host.GetHashCode()
                .ToString(CultureInfo.InvariantCulture);

            RegionManager.SetRegionManager(
                target: host,
                value: hostRegionManager);
            RegionManager.SetRegionName(
                regionTarget: host,
                regionName: result);
            RegionManager.UpdateRegions();

            var contentControl = taskPane.ContentControl as ContainerControl;
            var elementHost = contentControl.Controls[0] as ElementHost;

            elementHost.Child = host;

            return result;
        }

        private PrismTaskPanesView GetHostView(string hostRegion)
        {
            var horstUri = TaskPaneExtensions.GetUriHost();

            hostRegionManager.RequestNavigate(
                regionName: hostRegion,
                source: horstUri);

            var result = hostRegionManager.Regions[hostRegion].Views
                .SingleOrDefault() as PrismTaskPanesView;

            return result;
        }

        private CustomTaskPane GetTaskPane(TaskPaneSettings settings)
        {
            if (Type.GetTypeFromProgID(ProgID) == default)
            {
                throw new LibraryNotRegisteredException();
            }

            var result = ctpFactory.CreateCTP(
                cTPAxID: ProgID,
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

            result.VisibleStateChangeEvent += (taskPane) => BaseProvider.OnTaskPaneChanged(taskPane);

            return result;
        }

        private void SetRegions(PrismTaskPanesView hostView, TaskPaneSettings settings)
        {
            hostView.SetLocalRegion(
                regionName: settings.RegionName,
                regionContext: settings.RegionContext);

            var viewUri = settings.GetUriView(windowKey);
            var viewRegionManager = hostView.LocalRegionManager;

            viewRegionManager.RequestNavigate(
                regionName: settings.RegionName,
                source: viewUri);
        }

        #endregion Private Methods
    }
}

#pragma warning restore CA1031 // Keine allgemeinen Ausnahmetypen abfangen