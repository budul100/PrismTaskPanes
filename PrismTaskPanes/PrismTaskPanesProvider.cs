using CommonServiceLocator;
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
using System.Linq;
using System.Reflection;

namespace PrismTaskPanes
{
    public static class PrismTaskPanesProvider
    {
        #region Private Fields

        private const string TXTSettingsFileName = "PrismTaskPanes.xml";

        private static readonly List<IPrismTaskPaneReceiver> receivers =
            new List<IPrismTaskPaneReceiver>();

        private static object application;
        private static object ctpFactoryInst;
        private static IOfficeApplication officeApplication;

        #endregion Private Fields

        #region Public Properties

        public static EventHandler OnTaskPaneChanged { get; set; }

        public static IServiceLocator ServiceLocator => officeApplication?.ServiceLocator;

        #endregion Public Properties

        #region Public Methods

        public static void AddCTPReceiver(this IPrismTaskPaneReceiver receiver, object application, object ctpFactoryInst)
        {
            PrismTaskPanesProvider.application = application;
            PrismTaskPanesProvider.ctpFactoryInst = ctpFactoryInst;

            receivers.Add(receiver);
        }

        public static void ConfigureModuleCatalog(IModuleCatalog moduleCatalog)
        {
            foreach (var receiver in receivers)
            {
                receiver.ConfigureModuleCatalog(moduleCatalog);
            }
        }

        public static bool GetTaskPaneVisibility(this IPrismTaskPaneReceiver receiver, string id)
        {
            var hash = receiver.GetReceiverHash(id);
            var result = officeApplication?.GetTaskPaneVisibility(hash);

            return result ?? false;
        }

        public static void Initialize()
        {
            officeApplication = GetOfficeApplication();

            OnTaskPaneChanged += InvalidateReceivers;
        }

        public static void RegisterTypes(IContainerRegistry containerRegistry)
        {
            foreach (var receiver in receivers)
            {
                receiver.RegisterTypes(containerRegistry);
            }
        }

        public static void SetTaskPaneVisibility(this IPrismTaskPaneReceiver receiver, string id, bool isVisible)
        {
            var hash = HashExtensions.GetReceiverHash(
                id: id,
                receiver: receiver);
            officeApplication?.SetTaskPaneVisibility(
                hash: hash,
                isVisible: isVisible);
        }

        #endregion Public Methods

        #region Private Methods

        private static IEnumerable<PrismTaskPaneAttribute> GetAllAttributes()
        {
            foreach (var receiver in receivers)
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
        }

        private static IOfficeApplication GetOfficeApplication()
        {
            var result = officeApplication;

            if (result == default)
            {
                var settingsRepository = new SettingsRepository(
                    attributes: GetAllAttributes(),
                    settingsPath: GetSettingsPath());

                if (application is NetOffice.ExcelApi.Application)
                {
                    result = new ExcelApplication(
                        application: application,
                        ctpFactoryInst: ctpFactoryInst,
                        settingsRepository: settingsRepository);
                }
                // else if (application is Word.Application)
            }

            return result;
        }

        private static string GetSettingsPath()
        {
            var path = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                Assembly.GetExecutingAssembly().GetName().Name);
            var result = Path.Combine(path, TXTSettingsFileName);

            return result;
        }

        private static void InvalidateReceivers(object sender, EventArgs e)
        {
            if (receivers?.Any() ?? false)
            {
                foreach (var receiver in receivers)
                {
                    receiver.InvalidateRibbonUI();
                }
            }
        }

        #endregion Private Methods
    }
}