using DryIoc;
using NetOffice;
using NetOffice.OfficeApi;
using Prism.DryIoc;
using Prism.Ioc;
using Prism.Modularity;
using Prism.Regions;
using PrismTaskPanes.DryIoc.Factories;
using PrismTaskPanes.EventArgs;
using PrismTaskPanes.Exceptions;
using PrismTaskPanes.Extensions;
using PrismTaskPanes.Interfaces;
using System;
using System.Windows;

namespace PrismTaskPanes.DryIoc
{
    public abstract class BaseProvider
        : IDisposable
    {
        #region Private Fields

        private static bool alreadyLoadedMessageShown;
        private readonly object ctpFactoryInst;
        private readonly object officeApplication;

        private Application dryIocApplication;
        private bool isLoaded;
        private TaskPanesRepositoryFactory repositoryFactory;

        #endregion Private Fields

        #region Protected Constructors

        protected BaseProvider(ITaskPanesReceiver receiver, object officeApplication, object ctpFactoryInst)
        {
            Receiver = receiver ?? throw new ArgumentNullException(nameof(receiver));
            this.officeApplication = officeApplication ?? throw new ArgumentNullException(nameof(officeApplication));
            this.ctpFactoryInst = ctpFactoryInst ?? throw new ArgumentNullException(nameof(ctpFactoryInst));

            Service.AddReceiver(Receiver);
            Service.OnTaskPaneChangedEvent += OnTaskPaneChanged;
        }

        #endregion Protected Constructors

        #region Public Events

        public event EventHandler<ProviderEventArgs<IRegionBehaviorFactory>> OnConfigureDefaultRegionBehaviorsEvent;

        public event EventHandler<ProviderEventArgs<IModuleCatalog>> OnConfigureModuleCatalogEvent;

        public event EventHandler<ProviderEventArgs<IContainerRegistry>> OnRegisterTypesEvent;

        public event EventHandler<ProviderEventArgs<IResolverContext>> OnScopeClosingEvent;

        public event EventHandler<ProviderEventArgs<IResolverContext>> OnScopeOpenedEvent;

        public event EventHandler<TaskPaneEventArgs> OnTaskPaneChangedEvent;

        #endregion Public Events

        #region Public Properties

        public IResolverContext Scope => (repositoryFactory?.Get()?.Scope as IResolverContext)
            ?? (dryIocApplication?.Container?.GetContainer());

        #endregion Public Properties

        #region Protected Properties

        protected bool IsDisposed { get; private set; }

        protected ITaskPanesReceiver Receiver { get; }

        protected Func<string> TaskPaneIdentifierGetter { get; set; }

        protected Func<object> TaskPaneWindowGetter { get; set; }

        protected Func<int?> TaskPaneWindowKeyGetter { get; set; }

        #endregion Protected Properties

        #region Public Methods

        public virtual void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        public void InitilializeTaskPane(bool checkApplication = false)
        {
            if (checkApplication && !CheckApplication())
            {
                ShowAlreadyLoadedMessage();
            }
            else
            {
                StartApplication();
            }
        }

        public void SetTaskPaneVisibility(string id, bool isVisible,
            bool checkApplication = false)
        {
            if (checkApplication && !CheckApplication())
            {
                ShowAlreadyLoadedMessage();
            }
            else if (StartApplication())
            {
                var repository = repositoryFactory?.Get();

                if (repository != default)
                {
                    var hash = Receiver.GetReceiverHash(id);

                    repository.SetVisible(
                        receiverHash: hash,
                        isVisible: isVisible);
                }

                Receiver.InvalidateRibbonUI();
            }
        }

        public bool TaskPaneExists(string id, bool checkApplication = false)
        {
            var result = false;

            if (checkApplication && !CheckApplication())
            {
                ShowAlreadyLoadedMessage();
            }
            else
            {
                var repository = repositoryFactory?.Get();

                if (repository != default)
                {
                    var hash = Receiver.GetReceiverHash(id);

                    result = repository?.Exists(hash) ?? false;
                }
            }

            return result;
        }

        public bool TaskPaneIsVisible(string id, bool checkApplication = false)
        {
            var result = false;

            if (checkApplication && !CheckApplication())
            {
                ShowAlreadyLoadedMessage();
            }
            else
            {
                var repository = repositoryFactory?.Get();

                if (repository != default)
                {
                    var hash = Receiver.GetReceiverHash(id);

                    result = repository?.IsVisible(hash) ?? false;
                }
            }

            return result;
        }

        #endregion Public Methods

        #region Protected Methods

        protected virtual void Dispose(bool disposing)
        {
            if (!IsDisposed)
            {
                if (disposing)
                {
                    repositoryFactory?.Dispose();
                }

                IsDisposed = true;
            }
        }

        protected void OnApplicationDispose(OnDisposeEventArgs eventArgs)
        {
            if (!eventArgs.Cancel)
            {
                repositoryFactory?.Close();
            }
        }

        protected void OpenScope()
        {
            repositoryFactory?.Create();
        }

        protected void SaveScope()
        {
            var repository = repositoryFactory?.Get();

            repository?.Save();
        }

        #endregion Protected Methods

        #region Private Methods

        private static void ShowAlreadyLoadedMessage()
        {
            if (!alreadyLoadedMessageShown)
            {
                alreadyLoadedMessageShown = MessageBox.Show(
                    messageBoxText: $"{AlreadyLoadedException.Text}\r\n\r\nShall this message suppressed for this session?",
                    caption: "Already loaded task pane",
                    button: MessageBoxButton.YesNo,
                    icon: MessageBoxImage.Error,
                    defaultResult: MessageBoxResult.No) == MessageBoxResult.Yes;
            }
        }

        private bool CheckApplication()
        {
            var result = System.Windows.Application.Current == default || isLoaded;

            return result;
        }

        private void OnApplicationCreated(object sender, System.EventArgs e)
        {
            OpenScope();
        }

        private void OnConfigureDefaultRegionBehaviors(object sender, ProviderEventArgs<IRegionBehaviorFactory> e)
        {
            OnConfigureDefaultRegionBehaviorsEvent?.Invoke(
                sender: this,
                e: e);
        }

        private void OnConfigureModuleCatalog(object sender, ProviderEventArgs<IModuleCatalog> e)
        {
            OnConfigureModuleCatalogEvent?.Invoke(
                sender: this,
                e: e);
        }

        private void OnRegisterTypes(object sender, ProviderEventArgs<IContainerRegistry> e)
        {
            e.Content.RegisterInstance(officeApplication);
            e.Content.RegisterInstance(Receiver);
            e.Content.Register<IResolverContext>(() => Scope);

            OnRegisterTypesEvent?.Invoke(
                sender: this,
                e: e);
        }

        private void OnScopeClosing(object sender, ProviderEventArgs<IResolverContext> e)
        {
            Service.InvalidateRibbonUI();

            OnScopeClosingEvent?.Invoke(
                sender: this,
                e: e);
        }

        private void OnScopeOpened(object sender, ProviderEventArgs<IResolverContext> e)
        {
            Service.InvalidateRibbonUI();

            OnScopeOpenedEvent?.Invoke(
                sender: this,
                e: e);
        }

        private void OnTaskPaneChanged(object sender, TaskPaneEventArgs e)
        {
            OnTaskPaneChangedEvent?.Invoke(
                sender: this,
                e: e);
        }

        private bool StartApplication()
        {
            if (!CheckApplication())
            {
                throw new AlreadyLoadedException();
            }

            if (dryIocApplication == default)
            {
                dryIocApplication = System.Windows.Application.Current as Application
                    ?? new Application();

                dryIocApplication.OnConfigureDefaultRegionBehaviorsEvent += OnConfigureDefaultRegionBehaviors;
                dryIocApplication.OnConfigureModuleCatalogEvent += OnConfigureModuleCatalog;
                dryIocApplication.OnRegisterTypesEvent += OnRegisterTypes;
                dryIocApplication.OnApplicationCreatedEvent += OnApplicationCreated;

                var ctpFactory = new ICTPFactory(
                    parentObject: officeApplication as ICOMObject,
                    comProxy: ctpFactoryInst);

                repositoryFactory = new TaskPanesRepositoryFactory(
                    ctpFactory: ctpFactory,
                    dryIocApplication: dryIocApplication,
                    taskPaneWindowGetter: TaskPaneWindowGetter,
                    taskPaneWindowKeyGetter: TaskPaneWindowKeyGetter,
                    taskPaneIdentifierGetter: TaskPaneIdentifierGetter);

                repositoryFactory.OnScopeOpenedEvent += OnScopeOpened;
                repositoryFactory.OnScopeClosingEvent += OnScopeClosing;

                // isLoaded must be used since dryIocApplication cannot be requested directly
                // If the check is dryIocApplication == System.Windows.Application.Current,
                // then the IoC is not working correctly anymore

                isLoaded = true;
            }

            return dryIocApplication != default;
        }

        #endregion Private Methods
    }
}