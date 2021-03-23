using DryIoc;
using PrismTaskPanes.DryIoc.EventArgs;
using PrismTaskPanes.EventArgs;
using PrismTaskPanes.Interfaces;
using System;

namespace PrismTaskPanes.DryIoc
{
    public static class PowerPointProvider
    {
        #region Public Events

        public static event EventHandler OnApplicationExitEvent;

        public static event EventHandler<PowerPointEventArgs> OnScopeClosingEvent;

        public static event EventHandler<PowerPointEventArgs> OnScopeInitializedEvent;

        public static event EventHandler<PowerPointEventArgs> OnScopeOpenedEvent;

        public static event EventHandler<TaskPaneEventArgs> OnTaskPaneChangedEvent;

        #endregion Public Events

        #region Public Properties

        public static IResolverContext Container => DryIocProvider.Container;

        #endregion Public Properties

        #region Public Methods

        public static void InitializeProvider(this ITaskPanesReceiver receiver, object application,
            object ctpFactoryInst)
        {
            if (receiver == default)
            {
                throw new ArgumentNullException(nameof(receiver));
            }

            BaseProvider.AddReceiver(receiver);

            if (DryIocProvider.Application == default)
            {
                var dryIocApplication = new Application.PowerPointApplication(
                    application: application,
                    ctpFactoryInst: ctpFactoryInst);

                DryIocProvider.InitializeApplication(dryIocApplication);

                DryIocProvider.OnApplicationExitEvent += OnApplicationExit;
                DryIocProvider.OnScopeClosingEvent += OnScopeClosing;
                DryIocProvider.OnScopeInitializedEvent += OnScopeInitialized;
                DryIocProvider.OnScopeOpenedEvent += OnScopeOpened;
                DryIocProvider.OnTaskPaneChangedEvent += OnTaskPaneChanged;
            }
        }

        public static void SetTaskPaneVisible(this ITaskPanesReceiver receiver, string id, bool isVisible)
        {
            DryIocProvider.SetTaskPaneVisible(
                receiver: receiver,
                id: id,
                isVisible: isVisible);
        }

        public static bool TaskPaneExists(this ITaskPanesReceiver receiver, string id)
        {
            return DryIocProvider.TaskPaneExists(
                receiver: receiver,
                id: id);
        }

        public static bool TaskPaneVisible(this ITaskPanesReceiver receiver, string id)
        {
            return DryIocProvider.TaskPaneVisible(
                receiver: receiver,
                id: id);
        }

        #endregion Public Methods

        #region Private Methods

        private static void OnApplicationExit(object sender, System.EventArgs e)
        {
            OnApplicationExitEvent?.Invoke(
                sender: sender,
                e: e);
        }

        private static void OnScopeClosing(object sender, EventArgs.DryIocEventArgs e)
        {
            var eventArgs = new PowerPointEventArgs(e.Container);

            OnScopeClosingEvent?.Invoke(
                sender: sender,
                e: eventArgs);
        }

        private static void OnScopeInitialized(object sender, EventArgs.DryIocEventArgs e)
        {
            var eventArgs = new PowerPointEventArgs(e.Container);

            OnScopeInitializedEvent?.Invoke(
                sender: sender,
                e: eventArgs);
        }

        private static void OnScopeOpened(object sender, EventArgs.DryIocEventArgs e)
        {
            var eventArgs = new PowerPointEventArgs(e.Container);

            OnScopeOpenedEvent?.Invoke(
                sender: sender,
                e: eventArgs);
        }

        private static void OnTaskPaneChanged(object sender, TaskPaneEventArgs e)
        {
            OnTaskPaneChangedEvent?.Invoke(
                sender: sender,
                e: e);
        }

        #endregion Private Methods
    }
}