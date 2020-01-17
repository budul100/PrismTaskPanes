using CommonServiceLocator;
using DryIoc;
using NetOffice.OfficeApi;
using Prism.DryIoc;
using Prism.Ioc;
using Prism.Modularity;
using Prism.Regions;
using PrismTaskPanes.Controls;
using PrismTaskPanes.Extensions;
using PrismTaskPanes.Interfaces;
using PrismTaskPanes.Regions;
using PrismTaskPanes.TaskPanes;
using System;
using System.Collections.Generic;
using System.Linq;

namespace PrismTaskPanes.Applications.DryIoc
{
    internal abstract class OfficeApplication
        : PrismApplication, IDisposable, IOfficeApplication
    {
        #region Private Fields

        private readonly ICTPFactory ctpFactory;
        private readonly SettingsRepository settingsRepository;
        private readonly IList<TaskPanesRepository> taskPaneRepositories = new List<TaskPanesRepository>();

        private bool disposed;

        #endregion Private Fields

        #region Public Constructors

        public OfficeApplication(object application, object ctpFactoryInst, SettingsRepository settingsRepository)
        {
            this.settingsRepository = settingsRepository;

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
            CloseScope();

            Dispose(true);
            GC.SuppressFinalize(this);
        }

        public virtual void Dispose(bool disposing)
        {
            if (!disposed)
            {
                if (disposing)
                {
                    foreach (var taskPaneRepository in taskPaneRepositories)
                    {
                        CloseRepository(taskPaneRepository);
                    }

                    taskPaneRepositories.Clear();

                    ctpFactory.Dispose();
                }

                disposed = true;
            }
        }

        public IScope GetCurrentScope()
        {
            return Container.GetContainer().GetNamedScope(
                name: TaskPaneWindowKey.Value,
                throwIfNotFound: false);
        }

        public void SetTaskPaneVisible(int hash, bool isVisible)
        {
            var repository = GetRepository();

            if (repository != default)
            {
                repository.SetVisible(
                    hash: hash,
                    isVisible: isVisible);

                TaskPanesProvider.InvalidateRibbonUI();
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

        protected void CloseScope()
        {
            var repository = GetRepository();
            CloseRepository(repository);
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

            TaskPanesProvider.ConfigureModuleCatalog(moduleCatalog);
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

            TaskPanesProvider.RegisterTypes(containerRegistry);
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
                taskPaneRepositories.Remove(repository);
                repository.Dispose();

                var scope = GetCurrentScope();
                scope?.Dispose();
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

                    result = new TaskPanesRepository(
                        key: TaskPaneWindowKey.Value,
                        scope: scope,
                        taskPanesFactory: taskPanesFactory,
                        settingsRepository: settingsRepository,
                        documentHashGetter: () => GetDocumentHash());

                    taskPaneRepositories.Add(result);

                    result.Initialise();

                    TaskPanesProvider.OnTaskPaneInitializedEvent?.Invoke(
                        sender: scope,
                        e: null);
                }
            }

            return result;
        }

        #endregion Private Methods
    }
}