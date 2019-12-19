using Prism.Ioc;
using Prism.Modularity;

namespace PrismTaskPanes.Interfaces
{
    public interface IPrismTaskPaneReceiver
    {
        #region Public Methods

        void ConfigureModuleCatalog(IModuleCatalog moduleCatalog);

        void InvalidateRibbonUI();

        void RegisterTypes(IContainerRegistry builder);

        #endregion Public Methods
    }
}