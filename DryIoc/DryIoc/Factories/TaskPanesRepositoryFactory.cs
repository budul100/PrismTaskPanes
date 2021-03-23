using DryIoc;
using NetOffice.OfficeApi;
using Prism.Ioc;
using Prism.Regions;
using PrismTaskPanes.Factories;
using PrismTaskPanes.Settings;
using System;
using System.Collections.Generic;

namespace PrismTaskPanes.DryIoc.Factories
{
    public class TaskPanesRepositoryFactory
        : IDisposable
    {
        #region Private Fields

        private readonly Func<IContainer> containerGetter;
        private readonly ICTPFactory ctpFactory;
        private readonly IDictionary<int, TaskPanesRepository> repositories = new Dictionary<int, TaskPanesRepository>();
        private readonly Func<string> taskPaneIdentifierGetter;
        private readonly Func<object> taskPaneWindowGetter;
        private readonly Func<int?> taskPaneWindowKeyGetter;

        private bool isDisposed;

        #endregion Private Fields

        #region Public Constructors

        public TaskPanesRepositoryFactory(object application, object ctpFactoryInst, Func<IContainer> containerGetter,
            Func<object> taskPaneWindowGetter, Func<int?> taskPaneWindowKeyGetter, Func<string> taskPaneIdentifierGetter)
        {
            this.containerGetter = containerGetter;
            this.taskPaneWindowGetter = taskPaneWindowGetter;
            this.taskPaneWindowKeyGetter = taskPaneWindowKeyGetter;
            this.taskPaneIdentifierGetter = taskPaneIdentifierGetter;

            ctpFactory = new ICTPFactory(
                parentObject: application as NetOffice.ICOMObject,
                comProxy: ctpFactoryInst);
        }

        #endregion Public Constructors

        #region Public Methods

        public void Close()
        {
            CloseRepositories();
        }

        public void Create()
        {
            var result = Get();

            if (result == default)
            {
                var key = taskPaneWindowKeyGetter.Invoke();

                if (key.HasValue)
                {
                    CreateRepository(key.Value);
                }
            }
        }

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        public TaskPanesRepository Get()
        {
            var result = default(TaskPanesRepository);

            var key = taskPaneWindowKeyGetter.Invoke();

            if (key.HasValue
                && repositories.ContainsKey(key.Value))
            {
                result = repositories[key.Value];
            }

            return result;
        }

        public bool IsAvailable()
        {
            var key = taskPaneWindowKeyGetter.Invoke();

            var result = key.HasValue
                && repositories.ContainsKey(key.Value);

            return result;
        }

        #endregion Public Methods

        #region Protected Methods

        protected virtual void Dispose(bool disposing)
        {
            if (!isDisposed)
            {
                if (disposing)
                {
                    CloseRepositories();

                    containerGetter.Invoke().Dispose();
                    ctpFactory?.Dispose();
                }

                isDisposed = true;
            }
        }

        #endregion Protected Methods

        #region Private Methods

        private static void CloseRepository(TaskPanesRepository repository)
        {
            repository?.Save();
        }

        private void CloseRepositories()
        {
            foreach (var repository in repositories)
            {
                CloseRepository(repository.Value);
            }

            repositories.Clear();
        }

        private void CreateRepository(int key)
        {
            var scope = containerGetter.Invoke().OpenScope(key);

            DryIocProvider.OnScopeOpened(scope);

            var hostRegionManager = scope.Resolve<IRegionManager>();

            var taskPanesFactory = new TaskPanesFactory(
                key: key,
                ctpFactory: ctpFactory,
                hostRegionManager: hostRegionManager,
                taskPaneWindow: taskPaneWindowGetter.Invoke());

            var configurationsRepository = scope.Resolve<TaskPaneSettingsRepository>();

            var repository = new TaskPanesRepository(
                key: key,
                scope: scope,
                taskPanesFactory: taskPanesFactory,
                configurationsRepository: configurationsRepository,
                documentHashGetter: taskPaneIdentifierGetter);

            repositories.Add(
                key: key,
                value: repository);

            repository.OnRepositoryClosing += OnRepositoryClosing;

            repository.Initialise();

            DryIocProvider.OnScopeInitialized(scope);
        }

        private void OnRepositoryClosing(object sender, System.EventArgs e)
        {
            var repository = sender as TaskPanesRepository;

            DryIocProvider.OnScopeClosing(repository?.Scope as IResolverContext);
        }

        #endregion Private Methods
    }
}