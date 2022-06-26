using Prism.Mvvm;
using System.Runtime.InteropServices;

namespace Prism.Regions
{
    [ComVisible(false)]
    public abstract class ScopedContentBase
        : BindableBase, INavigationAware
    {
        #region Public Methods

        public virtual bool IsNavigationTarget(NavigationContext navigationContext)
        {
            return true;
        }

        public virtual void OnNavigatedFrom(NavigationContext navigationContext)
        { }

        public virtual void OnNavigatedTo(NavigationContext navigationContext)
        {
            Initialize(navigationContext);
        }

        #endregion Public Methods

        #region Protected Methods

        protected abstract void Initialize(NavigationContext navigationContext);

        #endregion Protected Methods
    }
}