using DryIoc;
using NetOffice.OfficeApi;
using Prism.Ioc;
using Prism.Regions;
using PrismTaskPanes.Settings;
using System;
using System.Collections.Generic;
using System.Linq;

namespace PrismTaskPanes.Factories
{
    internal class TaskPanesRepositoryFactory
        : IDisposable
    {
        #region Private Fields

        private readonly ICTPFactory ctpFactory;
        private readonly IList<TaskPanesRepository> repositories = new List<TaskPanesRepository>();
        private readonly Func<IContainer> scopeGetter;
        private readonly Func<string> taskPaneIdentifierGetter;
        private readonly Func<object> taskPaneWindowGetter;
        private readonly Func<int?> taskPaneWindowKeyGetter;

        private bool isDisposed = false;

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

            CloseRepository(repository);
        }

        public TaskPanesRepository Create()
        {
            var result = Get();

            if (result == default)
            {
                var key = taskPaneWindowKeyGetter.Invoke();

                if (key.HasValue)
                {
                    result = GetRepository(key.Value);
                }
            }

            return result;
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

            if (key.HasValue)
            {
                result = repositories
                    .SingleOrDefault(r => r.Key == key.Value);
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
                        CloseRepository(repository);
                    }

                    scopeGetter.Invoke().Dispose();
                    ctpFactory.Dispose();
                }

                repositories.Clear();

                isDisposed = true;
            }
        }

        #endregion Protected Methods

        #region Private Methods

        private void CloseRepository(TaskPanesRepository repository)
        {
            if (repository != default)
            {
                repository.Save();
                repository.Dispose();

                repositories.Remove(repository);
            }
        }

        private TaskPanesRepository GetRepository(int key)
        {
            var result = default(TaskPanesRepository);

            var scope = scopeGetter.Invoke().OpenScope(key);

            var hostRegionManager = scope.Resolve<IRegionManager>();

            var taskPanesFactory = new TaskPanesFactory(
                ctpFactory: ctpFactory,
                hostRegionManager: hostRegionManager,
                taskPaneWindow: taskPaneWindowGetter.Invoke());

            var configurationsRepository = scope.Resolve<TaskPaneSettingsRepository>();

            result = new TaskPanesRepository(
                key: key,
                scope: scope,
                taskPanesFactory: taskPanesFactory,
                configurationsRepository: configurationsRepository,
                documentHashGetter: taskPaneIdentifierGetter);

            repositories.Add(result);

            result.Initialise();

            DryIocProvider.OnTaskPaneInitializedEvent?.Invoke(
                sender: scope,
                e: default);

            return result;
        }

        #endregion Private Methods
    }
}