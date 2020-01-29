using DryIoc;
using Prism.Ioc;
using Prism.Modularity;
using PrismTaskPanes.Applications.DryIoc;
using PrismTaskPanes.Attributes;
using PrismTaskPanes.Configurations;
using PrismTaskPanes.Extensions;
using PrismTaskPanes.Interfaces;
using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;

namespace PrismTaskPanes
{
    public static class TaskPanesProvider
    {
        #region Private Fields

        private const string TXTSettingsFileName = "PrismTaskPanes.xml";

        private static readonly ConfigurationsRepository configurationsRepository = GetConfigurationsRepository();
        private static readonly IList<ITaskPanesReceiver> receivers = new List<ITaskPanesReceiver>();

        private static OfficeApplication officeApplication;

        #endregion Private Fields

        #region Public Properties

        public static EventHandler OnTaskPaneChangedEvent { get; set; }

        public static EventHandler OnTaskPaneInitializedEvent { get; set; }

        #endregion Public Properties

        #region Public Methods

        public static IScope GetCurrentScope()
        {
            return officeApplication.GetCurrentScope();
        }

        public static void InitializeTaskPanesProvider(this ITaskPanesReceiver receiver, object application,
            object ctpFactoryInst)
        {
            AddReceiver(receiver);

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

        #region Internal Methods

        internal static void ConfigureModuleCatalog(IModuleCatalog moduleCatalog)
        {
            DoWithAllReceivers((r) => r.ConfigureModuleCatalog(moduleCatalog));
        }

        internal static void InvalidateRibbonUI()
        {
            DoWithAllReceivers((r) => r.InvalidateRibbonUI());
        }

        internal static void RegisterTypes(IContainerRegistry containerRegistry)
        {
            containerRegistry.RegisterInstance(configurationsRepository);

            DoWithAllReceivers((r) => r.RegisterTypes(containerRegistry));
        }

        #endregion Internal Methods

        #region Private Methods

        private static void AddReceiver(ITaskPanesReceiver receiver)
        {
            var attributes = GetAttributes(receiver);
            configurationsRepository.AddAttributes(attributes);

            receivers.Add(receiver);
        }

        private static void DoWithAllReceivers(Action<ITaskPanesReceiver> setter)
        {
            foreach (var receiver in receivers)
            {
                setter?.Invoke(receiver);
            }
        }

        private static IEnumerable<PrismTaskPaneAttribute> GetAttributes(ITaskPanesReceiver receiver)
        {
            var type = receiver.GetType();

            var attributes = type.GetCustomAttributes(
                attributeType: typeof(PrismTaskPaneAttribute),
                inherit: true);

            foreach (var attribute in attributes)
            {
                var result = attribute as PrismTaskPaneAttribute;

                result.ReceiverHash = receiver.GetReceiverHash(result.ID);

                yield return result;
            }
        }

        private static ConfigurationsRepository GetConfigurationsRepository()
        {
            var directory = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                Assembly.GetExecutingAssembly().GetName().Name);
            var path = Path.Combine(directory, TXTSettingsFileName);

            var result = new ConfigurationsRepository(path);
            return result;
        }

        #endregion Private Methods
    }
}