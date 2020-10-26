using DryIoc;
using Prism.Ioc;
using Prism.Modularity;
using Prism.Regions;
using TestViewLib.Views;

namespace TestViewLib
{
    public class Module : IModule
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
            //containerRegistry.Register<ViewA>(typeof(ViewA).Name);
            //containerRegistry.GetContainer().Register<ViewA>(serviceKey: typeof(ViewA).Name, reuse: Reuse.Scoped);
        }

        #endregion Public Methods
    }
}