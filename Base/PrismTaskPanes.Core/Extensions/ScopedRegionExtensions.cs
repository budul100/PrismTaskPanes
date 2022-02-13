using Prism.Regions;
using PrismTaskPanes.Attributes;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows;

namespace PrismTaskPanes.Extensions
{
    [ComVisible(false)]
    public static class ScopedRegionExtensions
    {
        #region Public Methods

        public static ScopedRegionAttribute GetScopedRegionInfo(this object view)
        {
            return (ScopedRegionAttribute)view
                .GetType()
                .GetCustomAttributes(typeof(ScopedRegionAttribute), false).FirstOrDefault();
        }

        public static bool IsCandidate(this object view, NavigationContext navigationContext)
        {
            if ((view is INavigationAware navigationAware)
                && !navigationAware.IsNavigationTarget(navigationContext))
            {
                return false;
            }

            if (!(view is FrameworkElement frameworkElement))
            {
                return true;
            }

            navigationAware = frameworkElement.DataContext as INavigationAware;

            return navigationAware == default
                || navigationAware.IsNavigationTarget(navigationContext);
        }

        #endregion Public Methods
    }
}