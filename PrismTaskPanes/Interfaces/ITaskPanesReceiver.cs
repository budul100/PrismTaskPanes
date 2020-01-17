using Prism.Ioc;
using Prism.Modularity;

namespace PrismTaskPanes.Interfaces
{
    public interface ITaskPanesReceiver
    {
        #region Public Methods

        void ConfigureModuleCatalog(IModuleCatalog moduleCatalog);

        void InvalidateRibbonUI();

        void RegisterTypes(IContainerProvider containerProvider);

        #endregion Public Methods
    }
}