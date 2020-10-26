using DryIoc;
using Prism.DryIoc;
using Prism.Ioc;
using Prism.Modularity;
using Prism.Regions;
using PrismTaskPanes.Controls;
using PrismTaskPanes.Extensions;
using PrismTaskPanes.Factories;
using PrismTaskPanes.Regions;
using System;

namespace PrismTaskPanes.Applications.DryIoc
{
    internal abstract class OfficeApplication
        : PrismApplication, IDisposable
    {
        #region Protected Fields

        protected bool isDisposed;

        #endregion Protected Fields

        #region Private Fields

        private readonly TaskPanesRepositoryFactory repositoryFactory;
        private IResolverContext currentContainer;

        #endregion Private Fields

        #region Public Constructors

        public OfficeApplication(object application, object ctpFactoryInst)
        {
            DryIocProvider.OnScopeOpenedEvent += OnScopeOpened;
            DryIocProvider.OnTaskPaneChangedEvent += OnTaskPaneChanged;

            repositoryFactory = new TaskPanesRepositoryFactory(
                application: application,
                ctpFactoryInst: ctpFactoryInst,
                containerGetter: () => Container.GetContainer(),
                taskPaneWindowGetter: () => TaskPaneWindow,
                taskPaneWindowKeyGetter: () => TaskPaneWindowKey,
                taskPaneIdentifierGetter: () => GetTaskPaneIdentifier().GetHashString());
        }

        #endregion Public Constructors

        #region Protected Properties

        protected abstract object TaskPaneWindow { get; }

        protected abstract int? TaskPaneWindowKey { get; }

        #endregion Protected Properties

        #region Public Methods

        public virtual void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        public IResolverContext GetResolverContext()
        {
            var result = repositoryFactory.IsAvailable()
                ? repositoryFactory.Get().Scope as IResolverContext
                : currentContainer;

            return result;
        }

        public void SetTaskPaneVisible(string hash, bool isVisible)
        {
            var repository = repositoryFactory.Get();

            if (repository != default)
            {
                repository.SetVisible(
                    receiverHash: hash,
                    isVisible: isVisible);

                BaseProvider.InvalidateRibbonUI();
            }
        }

        public bool TaskPaneExists(string hash)
        {
            var result = repositoryFactory.Get()?
                .Exists(hash);

            return result ?? false;
        }

        public bool TaskPaneVisible(string hash)
        {
            var result = repositoryFactory.Get()?
                .IsVisible(hash);

            return result ?? false;
        }

        #endregion Public Methods

        #region Protected Methods

        protected override void ConfigureDefaultRegionBehaviors(IRegionBehaviorFactory regionBehaviors)
        {
            base.ConfigureDefaultRegionBehaviors(regionBehaviors);

            regionBehaviors.AddIfMissing(
                behaviorKey: DisposableRegionBehavior.BehaviorKey,
                behaviorType: typeof(DisposableRegionBehavior));
        }

        protected override void ConfigureModuleCatalog(IModuleCatalog moduleCatalog)
        {
            base.ConfigureModuleCatalog(moduleCatalog);

            BaseProvider.ConfigureModuleCatalog(moduleCatalog);
        }

        protected virtual void Dispose(bool disposing)
        {
            if (!isDisposed)
            {
                if (disposing)
                {
                    repositoryFactory.Dispose();
                }

                isDisposed = true;
            }
        }

        protected abstract string GetTaskPaneIdentifier();

        protected virtual void OnApplicationDispose(NetOffice.OnDisposeEventArgs eventArgs)
        {
            repositoryFactory.Close();
        }

        protected void OpenScope()
        {
            repositoryFactory.Create();
        }

        protected override void RegisterRequiredTypes(IContainerRegistry containerRegistry)
        {
            base.RegisterRequiredTypes(containerRegistry);

            containerRegistry.RegisterSingleton<IRegionNavigationContentLoader, ScopedRegionLoader>();
        }

        protected override void RegisterTypes(IContainerRegistry containerRegistry)
        {
            containerRegistry.Register<object, PrismTaskPanesView>(typeof(PrismTaskPanesView).FullName);
            containerRegistry.RegisterForNavigation<PrismTaskPanesView>();

            BaseProvider.RegisterTypes(containerRegistry);
        }

        protected void SaveScope()
        {
            var repository = repositoryFactory.Get();

            repository?.Save();
        }

        #endregion Protected Methods

        #region Private Methods

        private void OnScopeOpened(object sender, Events.DryIocEventArgs e)
        {
            currentContainer = e.Container;
        }

        private void OnTaskPaneChanged(object sender, Events.TaskPaneEventArgs e)
        {
            BaseProvider.InvalidateRibbonUI();
        }

        #endregion Private Methods
    }
}