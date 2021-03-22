using Prism.Ioc;
using Prism.Modularity;

namespace PrismTaskPanes.Interfaces
{
    public interface ITaskPanesReceiver
    {
        #region Public Methods

        void ConfigureModuleCatalog(IModuleCatalog moduleCatalog);

        void InvalidateRibbonUI();

        void RegisterTypes(IContainerRegistry containerRegistry);

        #endregion Public Methods
    }
}