using Prism.Regions;
using PrismTaskPanes.Attributes;
using PrismTaskPanes.Enums;
using PrismTaskPanes.Interfaces;
using System;
using System.Runtime.InteropServices;
using System.Windows.Controls;
using System.Windows.Data;

namespace PrismTaskPanes.Controls
{
    [ComVisible(false)]
    [ScopedRegion(
        CreateRegionManagerScope = true,
        ViewName = TaskPaneViewName)]
    public partial class PrismTaskPanesView
        : UserControl, IScopedRegion
    {
        #region Private Fields

        private const string LocalRegionManagerName = nameof(LocalRegionManager);
        private const string TaskPaneViewName = nameof(PrismTaskPanesView);

        private readonly ScrollViewer viewer;

        #endregion Private Fields

        #region Public Constructors

        public PrismTaskPanesView()
        {
            InitializeComponent();

            // Bind RegionManager.RegionManager attached property to this object
            var binding = new Binding(LocalRegionManagerName)
            {
                Mode = BindingMode.OneWayToSource,
                Source = this
            };

            SetBinding(
                dp: RegionManager.RegionManagerProperty,
                binding: binding);

            viewer = Content as ScrollViewer;
        }

        #endregion Public Constructors

        #region Public Properties

        public bool CreateRegionManagerScope => true;

        public IRegionManager LocalRegionManager { get; set; }

        public string ViewName => TaskPaneViewName;

        #endregion Public Properties

        #region Public Methods

        public static Uri GetHostUri()
        {
            var view = typeof(PrismTaskPanesView).Name;

            var result = new Uri(
                uriString: view,
                uriKind: UriKind.Relative);

            return result;
        }

        public void Initialize(string regionName, object regionContext, Uri viewUri,
            ScrollVisibility scrollBarHorizontal, ScrollVisibility scrollBarVertical)
        {
            SetLocalRegion(
                regionName: regionName,
                regionContext: regionContext);

            LocalRegionManager.RequestNavigate(
                regionName: regionName,
                source: viewUri);

            SetScrollBarHorizontal(scrollBarHorizontal);
            SetScrollBarVertical(scrollBarVertical);
        }

        #endregion Public Methods

        #region Private Methods

        private void SetLocalRegion(string regionName, object regionContext)
        {
            RegionManager.SetRegionManager(
                target: viewer,
                value: LocalRegionManager);
            RegionManager.SetRegionName(
                regionTarget: viewer,
                regionName: regionName);
            RegionManager.UpdateRegions();

            if (regionContext != default)
            {
                RegionManager.SetRegionContext(
                    target: viewer,
                    value: regionContext);
            }
        }

        private void SetScrollBarHorizontal(ScrollVisibility scrollBarVisibility)
        {
            switch (scrollBarVisibility)
            {
                case ScrollVisibility.Disabled:
                    viewer.HorizontalScrollBarVisibility = ScrollBarVisibility.Disabled;
                    break;

                case ScrollVisibility.Hidden:
                    viewer.HorizontalScrollBarVisibility = ScrollBarVisibility.Hidden;
                    break;

                case ScrollVisibility.Visible:
                    viewer.HorizontalScrollBarVisibility = ScrollBarVisibility.Visible;
                    break;

                default:
                    viewer.HorizontalScrollBarVisibility = ScrollBarVisibility.Auto;
                    break;
            }
        }

        private void SetScrollBarVertical(ScrollVisibility scrollBarVisibility)
        {
            switch (scrollBarVisibility)
            {
                case ScrollVisibility.Disabled:
                    viewer.VerticalScrollBarVisibility = ScrollBarVisibility.Disabled;
                    break;

                case ScrollVisibility.Hidden:
                    viewer.VerticalScrollBarVisibility = ScrollBarVisibility.Hidden;
                    break;

                case ScrollVisibility.Visible:
                    viewer.VerticalScrollBarVisibility = ScrollBarVisibility.Visible;
                    break;

                default:
                    viewer.VerticalScrollBarVisibility = ScrollBarVisibility.Auto;
                    break;
            }
        }

        #endregion Private Methods
    }
}