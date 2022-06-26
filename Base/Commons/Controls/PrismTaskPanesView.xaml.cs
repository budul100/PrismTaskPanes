using Prism.Regions;
using PrismTaskPanes.Attributes;
using PrismTaskPanes.Enums;
using PrismTaskPanes.Extensions;
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
        : UserControl, IScopedRegion, IPrismTaskPanesView
    {
        #region Private Fields

        private const string LocalRegionManagerName = nameof(LocalRegionManager);
        private const string TaskPaneViewName = nameof(PrismTaskPanesView);

        private readonly ScrollViewer viewer;

        #endregion Private Fields

        #region Public Constructors

        public PrismTaskPanesView()
        {
            this.LoadViewFromUri("/PrismTaskPanes.Commons;component/controls/prismtaskpanesview.xaml");

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
            viewer.HorizontalScrollBarVisibility = scrollBarVisibility switch
            {
                ScrollVisibility.Disabled => ScrollBarVisibility.Disabled,
                ScrollVisibility.Hidden => ScrollBarVisibility.Hidden,
                ScrollVisibility.Visible => ScrollBarVisibility.Visible,
                _ => ScrollBarVisibility.Auto,
            };
        }

        private void SetScrollBarVertical(ScrollVisibility scrollBarVisibility)
        {
            viewer.VerticalScrollBarVisibility = scrollBarVisibility switch
            {
                ScrollVisibility.Disabled => ScrollBarVisibility.Disabled,
                ScrollVisibility.Hidden => ScrollBarVisibility.Hidden,
                ScrollVisibility.Visible => ScrollBarVisibility.Visible,
                _ => ScrollBarVisibility.Auto,
            };
        }

        #endregion Private Methods
    }
}