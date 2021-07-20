using DryIoc;
using ExampleView.Views;
using Prism.Ioc;
using Prism.Modularity;
using Prism.Regions;

namespace ExampleView
{
    public class ExampleModule
        : IModule
    {
        #region Private Fields

        private const string TXTRegionName = "ExampleRegion";

        #endregion Private Fields

        #region Public Constructors

        public ExampleModule()
        {
            DummyClass.Dummy<Microsoft.Xaml.Behaviors.EventObserver>();
        }

        #endregion Public Constructors

        #region Public Methods

        public void OnInitialized(IContainerProvider containerProvider)
        {
            _ = containerProvider.Resolve<IRegionManager>().RegisterViewWithRegion(
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