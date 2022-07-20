using NetOffice.OfficeApi;
using NetOffice.OfficeApi.Enums;
using Prism.Regions;
using PrismTaskPanes.Controls;
using PrismTaskPanes.Exceptions;
using PrismTaskPanes.Extensions;
using PrismTaskPanes.Interfaces;
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

        private readonly ICTPFactory ctpFactory;
        private readonly IRegionManager hostRegionManager;
        private readonly string progId;
        private readonly object taskPaneWindow;
        private readonly int windowKey;

        #endregion Private Fields

        #region Public Constructors

        public TaskPanesFactory(ICTPFactory ctpFactory, IRegionManager hostRegionManager, object taskPaneWindow,
            int windowKey)
        {
            this.windowKey = windowKey;
            this.ctpFactory = ctpFactory;
            this.hostRegionManager = hostRegionManager;
            this.taskPaneWindow = taskPaneWindow;

            progId = typeof(PrismTaskPanesHost).FullName;

            if (Type.GetTypeFromProgID(progId) == default)
            {
                throw new HostNotRegisteredException();
            }
        }

        #endregion Public Constructors

        #region Public Methods

        public CustomTaskPane Get(TaskPaneSettings settings)
        {
            var result = GetTaskPane(settings);

            var hostRegion = GetHostRegion(result);
            var hostView = GetHostView(hostRegion);
            var viewUri = settings.GetUriView(windowKey);

            hostView.Initialize(
                regionName: settings.RegionName,
                regionContext: settings.RegionContext,
                viewUri: viewUri,
                scrollBarHorizontal: settings.ScrollBarHorizontal,
                scrollBarVertical: settings.ScrollBarVertical);

            return result;
        }

        #endregion Public Methods

        #region Private Methods

        private static Uri GetHostUri()
        {
            var uriString = typeof(PrismTaskPanesView).FullName;

            var result = new Uri(
                uriString: uriString,
                uriKind: UriKind.Relative);

            return result;
        }

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

        private IPrismTaskPanesView GetHostView(string hostRegion)
        {
            hostRegionManager.RequestNavigate(
                regionName: hostRegion,
                source: GetHostUri());

            var result = hostRegionManager.Regions[hostRegion].Views
                .OfType<IPrismTaskPanesView>().SingleOrDefault();

            return result;
        }

        private CustomTaskPane GetTaskPane(TaskPaneSettings settings)
        {
            var tpAxId = typeof(PrismTaskPanesHost).FullName;

            var result = ctpFactory.CreateCTP(
                cTPAxID: tpAxId,
                cTPTitle: settings.Title,
                cTPParentWindow: taskPaneWindow) as CustomTaskPane;

            try
            {
                result.Visible = false;
                result.DockPosition = settings.GetDockPosition();
                result.DockPositionRestrict = settings.GetDockRestriction();

                if (result.DockPosition != MsoCTPDockPosition.msoCTPDockPositionLeft &&
                    result.DockPosition != MsoCTPDockPosition.msoCTPDockPositionRight)
                {
                    result.Height = settings.Height;
                }

                if (result.DockPosition != MsoCTPDockPosition.msoCTPDockPositionBottom &&
                    result.DockPosition != MsoCTPDockPosition.msoCTPDockPositionTop)
                {
                    result.Width = settings.Width;
                }

                result.DockPositionStateChangeEvent += (t) => Service.OnTaskPaneChanged(t);
                result.VisibleStateChangeEvent += (t) => Service.OnTaskPaneChanged(t);
            }
            catch
            { }

            return result;
        }

        #endregion Private Methods
    }
}