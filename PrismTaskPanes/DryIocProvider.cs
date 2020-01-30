using DryIoc;
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

        public static IResolverContext ResolverContext
        {
            get => officeApplication.GetResolverContext();
        }

        #endregion Public Properties

        #region Public Methods

        public static void InitializeTaskPanesProvider(this ITaskPanesReceiver receiver, object application,
            object ctpFactoryInst)
        {
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

        public static void SetTaskPaneVisible(this ITaskPanesReceiver receiver, string id, bool isVisible)
        {
            var receiverHash = receiver.GetReceiverHash(id);
            officeApplication?.SetTaskPaneVisible(
                hash: receiverHash,
                isVisible: isVisible);
        }

        public static bool TaskPaneExists(this ITaskPanesReceiver receiver, string id)
        {
            var receiverHash = receiver.GetReceiverHash(id);
            var result = officeApplication?.TaskPaneExists(receiverHash);

            return result ?? false;
        }

        public static bool TaskPaneVisible(this ITaskPanesReceiver receiver, string id)
        {
            var receiverHash = receiver.GetReceiverHash(id);
            var result = officeApplication?.TaskPaneVisible(receiverHash);

            return result ?? false;
        }

        #endregion Public Methods
    }
}