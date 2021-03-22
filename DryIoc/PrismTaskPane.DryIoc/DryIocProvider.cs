using DryIoc;
using PrismTaskPanes.DryIoc.EventArgs;
using PrismTaskPanes.EventArgs;

using PrismTaskPanes.Extensions;
using PrismTaskPanes.Interfaces;
using System;

namespace PrismTaskPanes
{
    public static class DryIocProvider
    {
        #region Private Fields

        private static DryIoc.DryIocApplication application;

        #endregion Private Fields

        #region Public Events

        public static event EventHandler OnApplicationExitEvent;

        public static event EventHandler<DryIocEventArgs> OnScopeClosingEvent;

        public static event EventHandler<DryIocEventArgs> OnScopeInitialized;

        public static event EventHandler<DryIocEventArgs> OnScopeOpenedEvent;

        public static event EventHandler<TaskPaneEventArgs> OnTaskPaneChangedEvent;

        #endregion Public Events

        #region Internal Events

        internal static event EventHandler<DryIocEventArgs> OnScopeProvided;

        #endregion Internal Events

        #region Public Properties

        public static IResolverContext Container => application?.GetContainer();

        #endregion Public Properties

        #region Public Methods

        public static void InitializeApplication(DryIoc.DryIocApplication dryIocApplication)
        {
            if (dryIocApplication == default)
            {
                throw new ArgumentNullException(nameof(dryIocApplication));
            }

            if (application == default)
            {
                application = dryIocApplication;
                BaseProvider.OnTaskPaneChangedEvent += OnTaskPaneChanged;
            }
        }

        public static void SetTaskPaneVisible(this ITaskPanesReceiver receiver, string id, bool isVisible)
        {
            if (receiver == default)
            {
                throw new ArgumentNullException(nameof(receiver));
            }

            var receiverHash = receiver.GetReceiverHash(id);

            application?.SetTaskPaneVisible(
                hash: receiverHash,
                isVisible: isVisible);
        }

        public static bool TaskPaneExists(this ITaskPanesReceiver receiver, string id)
        {
            if (receiver == default)
            {
                throw new ArgumentNullException(nameof(receiver));
            }

            var receiverHash = receiver.GetReceiverHash(id);

            var result = application?.TaskPaneExists(receiverHash);

            return result ?? false;
        }

        public static bool TaskPaneVisible(this ITaskPanesReceiver receiver, string id)
        {
            if (receiver == default)
            {
                throw new ArgumentNullException(nameof(receiver));
            }

            var receiverHash = receiver.GetReceiverHash(id);
            var result = application?.TaskPaneVisible(receiverHash);

            return result ?? false;
        }

        #endregion Public Methods

        #region Internal Methods

        internal static void OnApplicationExit()
        {
            OnApplicationExitEvent?.Invoke(
                sender: application,
                e: default);
        }

        internal static void OnRepositoryInitialized(IResolverContext scope)
        {
            var eventArgs = new DryIocEventArgs(scope);

            OnScopeInitialized?.Invoke(
                sender: application,
                e: eventArgs);
        }

        internal static void OnScopeClosing(IResolverContext scope)
        {
            var eventArgs = new DryIocEventArgs(scope);

            OnScopeClosingEvent?.Invoke(
                sender: application,
                e: eventArgs);
        }

        internal static void OnScopeOpened(IResolverContext scope)
        {
            var eventArgs = new DryIocEventArgs(scope);

            OnScopeProvided?.Invoke(
                sender: application,
                e: eventArgs);

            OnScopeOpenedEvent?.Invoke(
                sender: application,
                e: eventArgs);
        }

        internal static void OnTaskPaneChanged(object sender, TaskPaneEventArgs e)
        {
            OnTaskPaneChangedEvent?.Invoke(
                sender: application,
                e: e);
        }

        #endregion Internal Methods
    }
}