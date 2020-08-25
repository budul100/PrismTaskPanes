using DryIoc;
using NetOffice.OfficeApi;
using PrismTaskPanes.Applications.DryIoc;
using PrismTaskPanes.Extensions;
using PrismTaskPanes.Interfaces;
using System;

namespace PrismTaskPanes
{
    public static class DryIocProvider
    {
        #region Private Fields

        private static OfficeApplication officeApplication;

        #endregion Private Fields

        #region Public Properties

        public static EventHandler OnTaskPaneChangedEvent { get; set; }

        public static EventHandler OnTaskPaneInitializedEvent { get; set; }

        public static IResolverContext ResolverContext => officeApplication?.GetResolverContext();

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

            if (officeApplication == default)
            {
                if (application is NetOffice.ExcelApi.Application)
                {
                    officeApplication = new ExcelApplication(
                        application: application,
                        ctpFactoryInst: ctpFactoryInst);
                }
            }
        }

        public static void RepositoryIsInitialized(IResolverContext scope)
        {
            OnTaskPaneInitializedEvent?.Invoke(
                sender: scope,
                e: default);
        }

        public static void SetTaskPaneVisible(this ITaskPanesReceiver receiver, string id, bool isVisible)
        {
            if (receiver == default)
            {
                throw new ArgumentNullException(nameof(receiver));
            }

            var receiverHash = receiver.GetReceiverHash(id);

            officeApplication?.SetTaskPaneVisible(
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

            var result = officeApplication?.TaskPaneExists(receiverHash);

            return result ?? false;
        }

        public static void TaskPaneIsChanged(_CustomTaskPane taskPane)
        {
            OnTaskPaneChangedEvent?.Invoke(
                sender: taskPane,
                e: default);
        }

        public static bool TaskPaneVisible(this ITaskPanesReceiver receiver, string id)
        {
            if (receiver == default)
            {
                throw new ArgumentNullException(nameof(receiver));
            }

            var receiverHash = receiver.GetReceiverHash(id);
            var result = officeApplication?.TaskPaneVisible(receiverHash);

            return result ?? false;
        }

        #endregion Public Methods
    }
}