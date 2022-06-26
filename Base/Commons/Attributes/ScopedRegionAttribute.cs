using PrismTaskPanes.Interfaces;
using System;
using System.Runtime.InteropServices;

namespace PrismTaskPanes.Attributes
{
    /// <summary>
    /// Attribute used to indicate if scoped region manager needs to be
    /// instantiated and attached to created view.
    /// </summary>
    [AttributeUsage(validOn: AttributeTargets.Class, AllowMultiple = false, Inherited = true)]
    [ComVisible(false)]
    public sealed class ScopedRegionAttribute
        : Attribute, IScopedRegion
    {
        #region Public Constructors

        /// <summary>
        /// By default it requires Scoped RegionManager and no Name;
        /// </summary>
        public ScopedRegionAttribute()
        {
            CreateRegionManagerScope = true;
        }

        #endregion Public Constructors

        #region Public Properties

        /// <summary>
        /// Indicates if new scoped RegionManaged should be associated with the View
        /// </summary>
        public bool CreateRegionManagerScope { get; set; }

        /// <summary>
        /// Name of the View
        /// </summary>
        public string ViewName { get; set; }

        #endregion Public Properties
    }
}