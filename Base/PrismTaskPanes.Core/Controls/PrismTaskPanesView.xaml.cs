using Prism.Regions;
using PrismTaskPanes.Attributes;
using PrismTaskPanes.Interfaces;
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
        }

        #endregion Public Constructors

        #region Public Properties

        public bool CreateRegionManagerScope => true;

        public IRegionManager LocalRegionManager { get; set; }

        public string ViewName => TaskPaneViewName;

        #endregion Public Properties

        #region Public Methods

        public void SetLocalRegion(string regionName, object regionContext)
        {
            var viewer = Content as ScrollViewer;

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

        #endregion Public Methods
    }
}