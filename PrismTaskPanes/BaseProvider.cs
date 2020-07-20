using Prism.Ioc;
using Prism.Modularity;
using PrismTaskPanes.Attributes;
using PrismTaskPanes.Extensions;
using PrismTaskPanes.Interfaces;
using PrismTaskPanes.Settings;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;

namespace PrismTaskPanes
{
    internal static class BaseProvider
    {
        #region Private Fields

        private const string TXTSettingsFileName = "PrismTaskPanes.xml";

        private static readonly TaskPaneSettingsRepository configurationsRepository = GetConfigurationsRepository();
        private static readonly IList<ITaskPanesReceiver> receivers = new List<ITaskPanesReceiver>();

        #endregion Private Fields

        #region Internal Methods

        internal static void AddReceiver(ITaskPanesReceiver receiver)
        {
            var attributes = GetAttributes(receiver).ToArray();
            configurationsRepository.AddAttributes(attributes);

            receivers.Add(receiver);
        }

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

        private static TaskPaneSettingsRepository GetConfigurationsRepository()
        {
            var directory = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                Assembly.GetExecutingAssembly().GetName().Name);
            var path = Path.Combine(directory, TXTSettingsFileName);

            var result = new TaskPaneSettingsRepository(path);
            return result;
        }

        #endregion Private Methods
    }
}