using PrismTaskPanes;
using PrismTaskPanes.Interfaces;
using System;

namespace PrismTaskPane.DryIoc
{
    public static class PowerPointProvider
    {
        #region Private Fields

        private static PrismTaskPanes.DryIoc.PowerPoint.PowerPointApplication dryIocApplication;

        #endregion Private Fields

        #region Public Methods

        public static void InitializeProvider(this ITaskPanesReceiver receiver, object application,
            object ctpFactoryInst)
        {
            if (receiver == default)
            {
                throw new ArgumentNullException(nameof(receiver));
            }

            BaseProvider.AddReceiver(receiver);

            if (dryIocApplication == default)
            {
                dryIocApplication = new PrismTaskPanes.DryIoc.PowerPoint.PowerPointApplication(
                    application: application,
                    ctpFactoryInst: ctpFactoryInst);

                DryIocProvider.InitializeApplication(dryIocApplication);
            }
        }

        #endregion Public Methods
    }
}