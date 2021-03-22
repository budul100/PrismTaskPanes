using DryIoc;
using ExampleView.Views;
using Prism.Ioc;
using Prism.Modularity;
using Prism.Regions;

namespace ExampleView
{
    public class Module
        : IModule
    {
        #region Private Fields

        private const string TXTRegionName = "ExampleRegion";

        #endregion Private Fields

        #region Public Methods

        public void OnInitialized(IContainerProvider containerProvider)
        {
            containerProvider.Resolve<IRegionManager>().RegisterViewWithRegion(
                regionName: TXTRegionName,
                viewType: typeof(ViewAView));
        }

        public void RegisterTypes(IContainerRegistry containerRegistry)
        {
            containerRegistry.RegisterForNavigation<Views.ViewAView>();
        }

        #endregion Public Methods
    }
}