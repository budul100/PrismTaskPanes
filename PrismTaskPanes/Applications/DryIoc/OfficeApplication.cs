using CommonServiceLocator;
using DryIoc;
using NetOffice.OfficeApi;
using Prism.DryIoc;
using Prism.Ioc;
using Prism.Modularity;
using Prism.Regions;
using PrismTaskPanes.Configurations;
using PrismTaskPanes.Controls;
using PrismTaskPanes.Extensions;
using PrismTaskPanes.Regions;
using PrismTaskPanes.TaskPanes;
using System;
using System.Collections.Generic;
using System.Linq;

namespace PrismTaskPanes.Applications.DryIoc
{
    internal abstract class OfficeApplication
        : PrismApplication, IDisposable
    {
        #region Private Fields

        private readonly ICTPFactory ctpFactory;
        private readonly IList<TaskPanesRepository> taskPaneRepositories = new List<TaskPanesRepository>();

        private bool disposed;

        #endregion Private Fields

        #region Public Constructors

        public OfficeApplication(object application, object ctpFactoryInst)
        {
            ctpFactory = new ICTPFactory(
                parentObject: application as NetOffice.ICOMObject,
                comProxy: ctpFactoryInst);
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

        public IScope GetCurrentScope()
        {
            var result = default(IScope);

            if (TaskPaneWindowKey.HasValue)
            {
                result = Container.GetContainer().GetNamedScope(
                    name: TaskPaneWindowKey.Value,
                    throwIfNotFound: false);
            }

            return result;
        }

        public void SetTaskPaneVisible(int hash, bool isVisible)
        {
            var repository = GetRepository();

            if (repository != default)
            {
                repository.SetVisible(
                    receiverHash: hash,
                    isVisible: isVisible);

                BaseProvider.InvalidateRibbonUI();
            }
        }

        public bool TaskPaneExists(int hash)
        {
            var result = GetRepository()?.Exists(hash);

            return result ?? false;
        }

        public bool TaskPaneVisible(int hash)
        {
            var result = GetRepository()?
                .IsVisible(hash);

            return result ?? false;
        }

        #endregion Public Methods

        #region Protected Methods

        protected void CloseApplication()
        {
            foreach (var taskPaneRepository in taskPaneRepositories)
            {
                CloseRepository(taskPaneRepository);
            }

            taskPaneRepositories.Clear();

            ctpFactory.Dispose();
            Container.GetContainer().Dispose();
        }

        protected void CloseScope()
        {
            var repository = GetRepository();
            CloseRepository(repository);

            var scope = GetCurrentScope();
            scope?.Dispose();
        }

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
            if (!disposed)
            {
                if (disposing)
                {
                    CloseApplication();
                }

                disposed = true;
            }
        }

        protected abstract string GetTaskPaneIdentifier();

        protected void OpenScope()
        {
            GetRepository(true);
        }

        protected override void RegisterRequiredTypes(IContainerRegistry containerRegistry)
        {
            base.RegisterRequiredTypes(containerRegistry);

            containerRegistry.RegisterSingleton<IServiceLocator, DryIocServiceLocatorAdapter>();
            containerRegistry.RegisterSingleton<IRegionNavigationContentLoader, ScopedRegionLoader>();
        }

        protected override void RegisterTypes(IContainerRegistry containerRegistry)
        {
            containerRegistry.Register<object, PrismTaskPanesView>(typeof(PrismTaskPanesView).Name);
            containerRegistry.RegisterForNavigation<PrismTaskPanesView>();

            BaseProvider.RegisterTypes(containerRegistry);
        }

        protected void SaveScope()
        {
            var repository = GetRepository();
            repository.Save();
        }

        #endregion Protected Methods

        #region Private Methods

        private void CloseRepository(TaskPanesRepository repository)
        {
            if (repository != default)
            {
                repository.Save();
                repository.Dispose();

                taskPaneRepositories.Remove(repository);
            }
        }

        private int GetDocumentHash()
        {
            return GetTaskPaneIdentifier()
                .GetStaticHash();
        }

        private TaskPanesRepository GetRepository(bool createIfNotExists = false)
        {
            var result = default(TaskPanesRepository);

            if (TaskPaneWindowKey.HasValue)
            {
                result = taskPaneRepositories
                    .SingleOrDefault(r => r.Key == TaskPaneWindowKey.Value);

                if (result == default && createIfNotExists)
                {
                    var scope = Container.GetContainer()
                        .OpenScope(name: TaskPaneWindowKey.Value);

                    var hostRegionManager = scope.Resolve<IRegionManager>();

                    var taskPanesFactory = new TaskPanesFactory(
                        ctpFactory: ctpFactory,
                        hostRegionManager: hostRegionManager,
                        taskPaneWindow: TaskPaneWindow);

                    var configurationsRepository = scope.Resolve<ConfigurationsRepository>();

                    result = new TaskPanesRepository(
                        key: TaskPaneWindowKey.Value,
                        scope: scope,
                        taskPanesFactory: taskPanesFactory,
                        configurationsRepository: configurationsRepository,
                        documentHashGetter: () => GetDocumentHash());

                    taskPaneRepositories.Add(result);

                    result.Initialise();

                    DryIocProvider.OnTaskPaneInitializedEvent?.Invoke(
                        sender: scope,
                        e: null);
                }
            }

            return result;
        }

        #endregion Private Methods
    }
}