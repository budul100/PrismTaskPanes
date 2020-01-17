using Prism.Ioc;
using Prism.Modularity;
using PrismTaskPanes.Applications.DryIoc;
using PrismTaskPanes.Attributes;
using PrismTaskPanes.Extensions;
using PrismTaskPanes.Interfaces;
using PrismTaskPanes.TaskPanes;
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

        private static OfficeApplication officeApplication;

        #endregion Private Fields

        #region Public Properties

        public static EventHandler OnTaskPaneChangedEvent { get; set; }

        public static EventHandler OnTaskPaneInitializedEvent { get; set; }

        #endregion Public Properties

        #region Public Methods

        public static void InitializeTaskPanesProvider(this ITaskPanesReceiver receiver, object application,
            object ctpFactoryInst)
        {
            if (officeApplication == default)
            {
                var settingsRepository = new SettingsRepository(
                    attributes: GetAllAttributes(receiver),
                    settingsPath: GetSettingsPath());

                if (application is NetOffice.ExcelApi.Application)
                {
                    officeApplication = new ExcelApplication(
                        taskPanesReceiver: taskPanesReceiver,
                        application: application,
                        ctpFactoryInst: ctpFactoryInst,
                        settingsRepository: settingsRepository);
                }
            }
        }

        public static void SetTaskPaneVisible(this ITaskPanesReceiver receiver, string id, bool isVisible)
        {
            var hash = receiver.GetReceiverHash(id);
            officeApplication?.SetTaskPaneVisible(
                hash: hash,
                isVisible: isVisible);
        }

        public static bool TaskPaneExists(this ITaskPanesReceiver receiver, string id)
        {
            var hash = receiver.GetReceiverHash(id);
            var result = officeApplication?.TaskPaneExists(hash);

            return result ?? false;
        }

        public static bool TaskPaneVisible(this ITaskPanesReceiver receiver, string id)
        {
            var hash = receiver.GetReceiverHash(id);
            var result = officeApplication?.TaskPaneVisible(hash);

            return result ?? false;
        }

        #endregion Public Methods

        #region Internal Methods

        internal static void ConfigureModuleCatalog(IModuleCatalog moduleCatalog)
        {
            taskPanesReceiver.ConfigureModuleCatalog(moduleCatalog);
        }

        internal static void InvalidateRibbonUI()
        {
            taskPanesReceiver.InvalidateRibbonUI();
        }

        internal static void RegisterTypes(IContainerRegistry containerRegistry)
        {
            taskPanesReceiver.RegisterTypes(containerRegistry);
        }

        #endregion Internal Methods

        #region Private Methods

        private static IEnumerable<PrismTaskPaneAttribute> GetAllAttributes(ITaskPanesReceiver receiver)
        {
            var type = receiver.GetType();

            var attributes = type.GetCustomAttributes(
                attributeType: typeof(PrismTaskPaneAttribute),
                inherit: true);

            foreach (var attribute in attributes)
            {
                var result = attribute as PrismTaskPaneAttribute;

                result.Hash = receiver.GetReceiverHash(result.ID);

                yield return result;
            }
        }

        private static string GetSettingsPath()
        {
            var path = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                Assembly.GetExecutingAssembly().GetName().Name);
            var result = Path.Combine(path, TXTSettingsFileName);

            return result;
        }

        #endregion Private Methods
    }
}