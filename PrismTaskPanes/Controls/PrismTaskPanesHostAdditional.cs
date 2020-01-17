using Prism.Regions;
using System.Runtime.InteropServices;
using System.Windows.Controls;

namespace PrismTaskPanes.Controls
{
    [ComVisible(true),
        ProgId("PrismTaskPanes.Controls.PrismTaskPanesHost"),
        Guid("6F14B2C2-3F59-456D-A224-50D76DF08176")]
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