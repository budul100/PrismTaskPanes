using Prism.Ioc;
using Prism.Modularity;
using System.Runtime.InteropServices;

namespace PrismTaskPanes.Interfaces
{
    [ComVisible(false)]
    public interface ITaskPanesReceiver
    {
        #region Public Methods

        void ConfigureModuleCatalog(IModuleCatalog moduleCatalog);

        void InvalidateRibbonUI();

        void RegisterTypes(IContainerRegistry containerRegistry);

        #endregion Public Methods
    }
}