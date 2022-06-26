using DryIoc;
using NetOffice.OfficeApi;
using Prism.Regions;
using PrismTaskPanes.EventArgs;
using PrismTaskPanes.Factories;
using System;
using System.Collections.Generic;

namespace PrismTaskPanes.DryIoc.Factories
{
    public class TaskPanesRepositoryFactory
        : IDisposable
    {
        #region Private Fields

        private readonly ICTPFactory ctpFactory;
        private readonly Application dryIocApplication;
        private readonly IDictionary<int, TaskPanesRepository> repositories = new Dictionary<int, TaskPanesRepository>();
        private readonly Func<string> taskPaneIdentifierGetter;
        private readonly Func<object> taskPaneWindowGetter;
        private readonly Func<int?> taskPaneWindowKeyGetter;

        private bool isDisposed;

        #endregion Private Fields

        #region Public Constructors

        public TaskPanesRepositoryFactory(ICTPFactory ctpFactory, Application dryIocApplication,
            Func<object> taskPaneWindowGetter, Func<int?> taskPaneWindowKeyGetter, Func<string> taskPaneIdentifierGetter)
        {
            this.ctpFactory = ctpFactory;
            this.dryIocApplication = dryIocApplication;
            this.taskPaneWindowGetter = taskPaneWindowGetter;
            this.taskPaneWindowKeyGetter = taskPaneWindowKeyGetter;
            this.taskPaneIdentifierGetter = taskPaneIdentifierGetter;
        }

        #endregion Public Constructors

        #region Public Events

        public event EventHandler<ProviderEventArgs<IResolverContext>> OnScopeClosingEvent;

        public event EventHandler<ProviderEventArgs<IResolverContext>> OnScopeOpenedEvent;

        #endregion Public Events

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
                    ctpFactory?.Dispose();
                }

                isDisposed = true;
            }
        }

        #endregion Protected Methods

        #region Private Methods

        private void CloseRepositories()
        {
            foreach (var repository in repositories)
            {
                CloseRepository(repository.Value);
            }

            repositories.Clear();
        }

        private void CloseRepository(TaskPanesRepository repository)
        {
            var scope = repository.Scope as IResolverContext;

            OnScopeClosing(scope);

            repository.Save();
            scope?.Dispose();
        }

        private void CreateRepository(int key)
        {
            var scope = dryIocApplication.OpenScope(key);

            var hostRegionManager = scope.Resolve<IRegionManager>();

            var taskPanesFactory = new TaskPanesFactory(
                ctpFactory: ctpFactory,
                hostRegionManager: hostRegionManager,
                taskPaneWindow: taskPaneWindowGetter.Invoke(),
                windowKey: key);

            var repository = new TaskPanesRepository(
                key: key,
                scope: scope,
                taskPanesFactory: taskPanesFactory,
                configurationsRepository: Service.ConfigurationsRepository,
                documentHashGetter: taskPaneIdentifierGetter);

            repositories.Add(
                key: key,
                value: repository);

            repository.Initialise();

            OnScopeOpened(scope);
        }

        private void OnScopeClosing(IResolverContext scope)
        {
            var eventArgs = new ProviderEventArgs<IResolverContext>(scope);

            OnScopeClosingEvent?.Invoke(
                sender: this,
                e: eventArgs);
        }

        private void OnScopeOpened(IResolverContext scope)
        {
            var eventArgs = new ProviderEventArgs<IResolverContext>(scope);

            OnScopeOpenedEvent?.Invoke(
                sender: this,
                e: eventArgs);
        }

        #endregion Private Methods
    }
}