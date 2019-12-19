using Prism.Ioc;
using Prism.Modularity;
using Prism.Regions;
using TestViewLib.ViewModels;
using TestViewLib.Views;

namespace TestViewLib
{
    public class TestViewLibModule : IModule
    {
        #region Private Fields

        private const string TXTRegionName = "TestRegion";

        #endregion Private Fields

        #region Public Methods

        public void OnInitialized(IContainerProvider containerProvider)
        {
            containerProvider.Resolve<IRegionManager>().RegisterViewWithRegion(
                regionName: TXTRegionName,
                viewType: typeof(ViewA));
        }

        public void RegisterTypes(IContainerRegistry containerRegistry)
        {
            containerRegistry.Register<ViewA>(typeof(ViewA).Name);
            //containerRegistry.Register<ViewAViewModel>();
        }

        #endregion Public Methods
    }
}