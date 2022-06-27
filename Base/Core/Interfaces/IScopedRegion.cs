using System.Runtime.InteropServices;

namespace PrismTaskPanes.Interfaces
{
    /// <summary>
    /// Interface used to indicate if scoped region manager needs to be
    /// instantiated and attached to created view.
    /// </summary>
    [ComVisible(false)]
    public interface IScopedRegion
    {
        #region Public Properties

        /// <summary>
        /// Indicates if new scoped RegionManaged should be associated with the View
        /// </summary>
        bool CreateRegionManagerScope { get; }

        /// <summary>
        /// Name of the View
        /// </summary>
        string ViewName { get; }

        #endregion Public Properties
    }
}