using DryIoc;
using Prism.DryIoc;
using Prism.Ioc;
using Prism.Modularity;
using Prism.Regions;
using PrismTaskPanes.Controls;
using PrismTaskPanes.DryIoc.EventArgs;
using PrismTaskPanes.DryIoc.Factories;
using PrismTaskPanes.Extensions;
using PrismTaskPanes.Regions;
using System;
using System.Windows;

namespace PrismTaskPanes.DryIoc.Application
{
    public abstract class DryIocApplication
        : PrismApplication, IDisposable
    {
        #region Protected Fields

        protected bool isDisposed;

        #endregion Protected Fields

        #region Private Fields

        private readonly Type contentType;
        private readonly TaskPanesRepositoryFactory repositoryFactory;

        #endregion Private Fields

        #region Protected Constructors

        protected DryIocApplication(object application, object ctpFactoryInst, Type contentType)
        {
            this.contentType = contentType;

            repositoryFactory = new TaskPanesRepositoryFactory(
                application: application,
                ctpFactoryInst: ctpFactoryInst,
                containerGetter: () => Container.GetContainer(),
                taskPaneWindowGetter: TaskPaneWindowGetter,
                taskPaneWindowKeyGetter: TaskPaneWindowKeyGetter,
                taskPaneIdentifierGetter: () => GetTaskPaneIdentifier().GetHashString());

            DryIocProvider.OnScopeOpenedEvent += OnScopeOpened;
            DryIocProvider.OnScopeClosingEvent += OnScopeClosing;
        }

        #endregion Protected Constructors

        #region Protected Properties

        protected abstract Func<object> TaskPaneWindowGetter { get; }

        protected abstract Func<int?> TaskPaneWindowKeyGetter { get; }

        #endregion Protected Properties

        #region Public Methods

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        public IResolverContext GetContainer()
        {
            var result = repositoryFactory.IsAvailable()
                ? repositoryFactory.Get().Scope as IResolverContext
                : Container.GetContainer();

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
            var repository = repositoryFactory.Get();

            var result = repository?.Exists(hash) ?? false;

            return result;
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

        protected override Window CreateShell()
        {
            DryIocProvider.OnProviderReady();

            return default;
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

        protected virtual void OnApplicationDispose()
        {
            repositoryFactory.Close();
        }

        protected override void OnExit(ExitEventArgs e)
        {
            DryIocProvider.OnApplicationExit();

            base.OnExit(e);
        }

        protected void OpenScope()
        {
            repositoryFactory.Create(contentType);
        }

        protected override void RegisterRequiredTypes(IContainerRegistry containerRegistry)
        {
            base.RegisterRequiredTypes(containerRegistry);

            containerRegistry.RegisterSingleton<IRegionNavigationContentLoader, ScopedRegionLoader>();
        }

        protected override void RegisterTypes(IContainerRegistry containerRegistry)
        {
            containerRegistry.Register<IResolverContext>(() => GetContainer());

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

        private void OnScopeClosing(object sender, DryIocEventArgs e)
        {
            BaseProvider.InvalidateRibbonUI();
        }

        private void OnScopeOpened(object sender, DryIocEventArgs e)
        {
            BaseProvider.InvalidateRibbonUI();
        }

        #endregion Private Methods
    }
}