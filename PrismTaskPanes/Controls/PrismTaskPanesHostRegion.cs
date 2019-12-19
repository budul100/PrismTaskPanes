using Prism.Regions;
using System.Runtime.InteropServices;
using System.Windows.Controls;

namespace PrismTaskPanes.Controls
{
    [ComVisible(true),
        ProgId("PrismTaskPanesCore.Controls.PrismTaskPanesHost"),
        Guid("1B17958C-545A-4AF2-A655-84709C46147C")]
    public partial class PrismTaskPanesHost
    {
        #region Public Methods

        public void SetLocalRegion(string regionName, IRegionManager regionManager)
        {
            var host = new ContentControl();

            RegionManager.SetRegionManager(
                target: host,
                value: regionManager);
            RegionManager.SetRegionName(
                regionTarget: host,
                regionName: regionName);
            RegionManager.UpdateRegions();

            elementHost.Child = host;
            elementHost.Dock = System.Windows.Forms.DockStyle.Fill;
        }

        #endregion Public Methods
    }
}