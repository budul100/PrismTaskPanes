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
using System.Windows;

namespace PrismTaskPanes.Applications.DryIoc
{
    internal abstract class OfficeApplication
        : PrismApplication, IDisposable, IOfficeApplication
    {
        #region Protected Fields

        protected readonly IList<TaskPanesRepository> taskPaneRepositories = new List<TaskPanesRepository>();
        protected bool isActivated;

        #endregion Protected Fields

        #region Private Fields

        private readonly SettingsRepository settingsRepository;
        private bool disposed = false;

        #endregion Private Fields

        #region Public Constructors

        public OfficeApplication(object application, object ctpFactoryInst, SettingsRepository settingsRepository)
        {
            this.settingsRepository = settingsRepository;

            CTPFactory = new ICTPFactory(
                parentObject: application as NetOffice.ICOMObject,
                comProxy: ctpFactoryInst);
        }

        #endregion Public Constructors

        #region Public Properties

        public IServiceLocator ServiceLocator { get; private set; }

        #endregion Public Properties

        #region Protected Properties

        protected ICTPFactory CTPFactory { get; }

        protected abstract object TaskPaneWindow { get; }

        protected abstract int? TaskPaneWindowKey { get; }

        #endregion Protected Properties

        #region Public Methods

        public virtual void Dispose()
        {
            CloseWindow();

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

                    CTPFactory.Dispose();
                }

                // TODO: nicht verwaltete Ressourcen (nicht verwaltete Objekte) freigeben und Finalizer weiter unten überschreiben.
                // TODO: große Felder auf Null setzen.

                disposed = true;
            }
        }

        public bool GetTaskPaneVisibility(int hash)
        {
            var result = GetRepository()?
                .GetVisibility(hash);

            return result ?? false;
        }

        public void SetTaskPaneVisibility(int hash, bool isVisible)
        {
            GetRepository()?.SetVisibility(
                hash: hash,
                isVisible: isVisible);
        }

        #endregion Public Methods

        #region Protected Methods

        protected void CloseWindow()
        {
            if (TaskPaneWindowKey.HasValue)
            {
                var respository = GetRepository(true);
                CloseRepository(respository);
            }
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

            PrismTaskPanesProvider.ConfigureModuleCatalog(moduleCatalog);
        }

        protected override Window CreateShell()
        {
            ServiceLocator = Container.Resolve<IServiceLocator>();

            return default;
        }

        protected TaskPanesRepository GetRepository(bool onlyExisting = false)
        {
            var result = default(TaskPanesRepository);

            if (TaskPaneWindowKey.HasValue)
            {
                result = taskPaneRepositories
                    .SingleOrDefault(r => r.Key == TaskPaneWindowKey.Value);

                if (result == default && !onlyExisting)
                {
                    var scope = Container.GetContainer().OpenScope();

                    var hostRegionManager = scope.Resolve<IRegionManager>();

                    var taskPanesFactory = new TaskPanesFactory(
                        ctpFactory: CTPFactory,
                        hostRegionManager: hostRegionManager,
                        taskPaneWindow: TaskPaneWindow);

                    result = new TaskPanesRepository(
                        key: TaskPaneWindowKey.Value,
                        taskPanesFactory: taskPanesFactory,
                        settingsRepository: settingsRepository,
                        documentHashGetter: () => GetDocumentHash());

                    taskPaneRepositories.Add(result);
                }
            }

            return result;
        }

        protected abstract string GetTaskPaneIdentifier();

        protected override void RegisterRequiredTypes(IContainerRegistry containerRegistry)
        {
            base.RegisterRequiredTypes(containerRegistry);

            containerRegistry
                .RegisterSingleton<IRegionNavigationContentLoader, ScopedRegionLoader>();
        }

        protected override void RegisterTypes(IContainerRegistry containerRegistry)
        {
            containerRegistry.Register<object, PrismTaskPanesView>(typeof(PrismTaskPanesView).Name);
            containerRegistry.RegisterForNavigation<PrismTaskPanesView>();

            PrismTaskPanesProvider.RegisterTypes(containerRegistry);
        }

        #endregion Protected Methods

        #region Private Methods

        private void CloseRepository(TaskPanesRepository respository)
        {
            if (respository != default)
            {
                if (isActivated) respository.Save();

                taskPaneRepositories.Remove(respository);
                respository.Dispose();
            }
        }

        private int GetDocumentHash()
        {
            return GetTaskPaneIdentifier()
                .GetStaticHash();
        }

        #endregion Private Methods
    }
}