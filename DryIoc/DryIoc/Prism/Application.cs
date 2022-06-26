using DryIoc;
using Prism.DryIoc;
using Prism.Ioc;
using Prism.Modularity;
using Prism.Regions;
using PrismTaskPanes.Controls;
using PrismTaskPanes.EventArgs;
using PrismTaskPanes.Regions;
using System;
using System.Windows;

namespace PrismTaskPanes.DryIoc
{
    public class Application
        : PrismApplication
    {
        #region Public Events

        public event EventHandler OnApplicationCreatedEvent;

        public event EventHandler OnApplicationExitEvent;

        public event EventHandler<ProviderEventArgs<IRegionBehaviorFactory>> OnConfigureDefaultRegionBehaviorsEvent;

        public event EventHandler<ProviderEventArgs<IModuleCatalog>> OnConfigureModuleCatalogEvent;

        public event EventHandler<ProviderEventArgs<IContainerRegistry>> OnRegisterTypesEvent;

        #endregion Public Events

        #region Public Methods

        public IResolverContext OpenScope(int key)
        {
            var result = Container.GetContainer()?.OpenScope(key);

            return result;
        }

        #endregion Public Methods

        #region Protected Methods

        protected override void ConfigureDefaultRegionBehaviors(IRegionBehaviorFactory regionBehaviors)
        {
            base.ConfigureDefaultRegionBehaviors(regionBehaviors);

            regionBehaviors.AddIfMissing(
                behaviorKey: DisposableRegionBehavior.BehaviorKey,
                behaviorType: typeof(DisposableRegionBehavior));

            var eventArgs = new ProviderEventArgs<IRegionBehaviorFactory>(regionBehaviors);

            OnConfigureDefaultRegionBehaviorsEvent?.Invoke(
                sender: this,
                e: eventArgs);
        }

        protected override void ConfigureModuleCatalog(IModuleCatalog moduleCatalog)
        {
            base.ConfigureModuleCatalog(moduleCatalog);

            var eventArgs = new ProviderEventArgs<IModuleCatalog>(moduleCatalog);

            OnConfigureModuleCatalogEvent?.Invoke(
                sender: this,
                e: eventArgs);
        }

        protected override Window CreateShell()
        {
            OnApplicationCreatedEvent?.Invoke(
                sender: this,
                e: default);

            return default;
        }

        protected override void OnExit(ExitEventArgs e)
        {
            OnApplicationExitEvent?.Invoke(
                sender: this,
                e: default);

            base.OnExit(e);
        }

        protected override void RegisterRequiredTypes(IContainerRegistry containerRegistry)
        {
            base.RegisterRequiredTypes(containerRegistry);

            containerRegistry.RegisterSingleton<IRegionNavigationContentLoader, ScopedRegionLoader>();
        }

        protected override void RegisterTypes(IContainerRegistry containerRegistry)
        {
            containerRegistry.Register<object, PrismTaskPanesView>(typeof(PrismTaskPanesView).FullName);
            containerRegistry.RegisterForNavigation<PrismTaskPanesView>(typeof(PrismTaskPanesView).FullName);

            var eventArgs = new ProviderEventArgs<IContainerRegistry>(containerRegistry);

            OnRegisterTypesEvent?.Invoke(
                sender: this,
                e: eventArgs);
        }

        #endregion Protected Methods
    }
}