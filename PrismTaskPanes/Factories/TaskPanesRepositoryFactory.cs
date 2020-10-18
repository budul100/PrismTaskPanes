using DryIoc;
using NetOffice.OfficeApi;
using Prism.Ioc;
using Prism.Regions;
using PrismTaskPanes.Settings;
using System;
using System.Collections.Generic;

namespace PrismTaskPanes.Factories
{
    internal class TaskPanesRepositoryFactory
        : IDisposable
    {
        #region Private Fields

        private readonly ICTPFactory ctpFactory;
        private readonly IDictionary<int, TaskPanesRepository> repositories = new Dictionary<int, TaskPanesRepository>();
        private readonly Func<IContainer> scopeGetter;
        private readonly Func<string> taskPaneIdentifierGetter;
        private readonly Func<object> taskPaneWindowGetter;
        private readonly Func<int?> taskPaneWindowKeyGetter;

        private bool isDisposed;

        #endregion Private Fields

        #region Public Constructors

        public TaskPanesRepositoryFactory(object application, object ctpFactoryInst, Func<IContainer> scopeGetter,
            Func<object> taskPaneWindowGetter, Func<int?> taskPaneWindowKeyGetter, Func<string> taskPaneIdentifierGetter)
        {
            this.scopeGetter = scopeGetter;
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
            var repository = Get();

            if (repository != default)
            {
                var key = repository.Key;

                CloseRepository(repository);
                repositories.Remove(key);
            }
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

        #endregion Public Methods

        #region Protected Methods

        protected virtual void Dispose(bool disposing)
        {
            if (!isDisposed)
            {
                if (disposing)
                {
                    foreach (var repository in repositories)
                    {
                        CloseRepository(repository.Value);
                    }

                    scopeGetter.Invoke().Dispose();
                    ctpFactory?.Dispose();
                }

                repositories.Clear();

                isDisposed = true;
            }
        }

        #endregion Protected Methods

        #region Private Methods

        private static void CloseRepository(TaskPanesRepository repository)
        {
            if (repository != default)
            {
                repository.Save();
                repository.Dispose();
            }
        }

        private void CreateRepository(int key)
        {
            var scope = scopeGetter.Invoke().OpenScope(key);

            var hostRegionManager = scope.Resolve<IRegionManager>();

            var taskPanesFactory = new TaskPanesFactory(
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

            repository.Initialise();
            DryIocProvider.RepositoryIsInitialized(scope);
        }

        #endregion Private Methods
    }
}