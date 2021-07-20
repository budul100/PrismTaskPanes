#pragma warning disable CA1031 // Do not catch general exception types

using NetOffice.OfficeApi;
using Prism.Ioc;
using Prism.Modularity;
using PrismTaskPanes.Controls;
using PrismTaskPanes.EventArgs;
using PrismTaskPanes.Extensions;
using PrismTaskPanes.Interfaces;
using PrismTaskPanes.Settings;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;

namespace PrismTaskPanes
{
    [ComVisible(false)]
    public static class BaseProvider
    {
        #region Private Fields

        private const string SettingsFile = "PrismTaskPanes.xml";

        private static readonly TaskPaneSettingsRepository configurationsRepository = GetConfigurationsRepository();
        private static readonly HashSet<ITaskPanesReceiver> receivers = new();

        #endregion Private Fields

        #region Public Events

        public static event EventHandler<TaskPaneEventArgs> OnTaskPaneChangedEvent;

        #endregion Public Events

        #region Public Methods

        public static void AddReceiver(ITaskPanesReceiver receiver)
        {
            var attributes = receiver
                .GetAttributes().ToArray();

            configurationsRepository.AddAttributes(attributes);

            _ = receivers.Add(receiver);
        }

        public static void ConfigureModuleCatalog(IModuleCatalog moduleCatalog)
        {
            DoWithAllReceivers((r) => r.ConfigureModuleCatalog(moduleCatalog));
        }

        public static void InvalidateRibbonUI()
        {
            DoWithAllReceivers((r) => r.InvalidateRibbonUI());
        }

        public static void RedirectAssembly()
        {
            AppDomain.CurrentDomain.AssemblyResolve += ResolveAssemblyOnCurrentDomain;
        }

        public static void RegisterProvider(Type contentType)
        {
            var progId = typeof(PrismTaskPanesHost).GetProgId(contentType);

            var directory = Path.Combine(
                path1: Environment.GetFolderPath(Environment.SpecialFolder.CommonProgramFiles),
                path2: nameof(PrismTaskPanesHost));

            ComExtensions.Register(
                progId: progId,
                type: typeof(PrismTaskPanesHost),
                directory: directory);
        }

        public static void RegisterTypes(IContainerRegistry containerRegistry)
        {
            _ = containerRegistry.RegisterInstance(configurationsRepository);

            DoWithAllReceivers((r) => r.RegisterTypes(containerRegistry));
        }

        public static void UnregisterProvider(Type contentType)
        {
            var progId = typeof(PrismTaskPanesHost).GetProgId(contentType);

            ComExtensions.Unregister(
                progId: progId);
        }

        #endregion Public Methods

        #region Internal Methods

        internal static void OnTaskPaneChanged(_CustomTaskPane taskPane)
        {
            InvalidateRibbonUI();

            var eventArgs = new TaskPaneEventArgs(taskPane);

            OnTaskPaneChangedEvent?.Invoke(
                sender: default,
                e: eventArgs);
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

        private static TaskPaneSettingsRepository GetConfigurationsRepository()
        {
            var directory = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                Assembly.GetExecutingAssembly().GetName().Name);
            var path = Path.Combine(directory, SettingsFile);

            var result = new TaskPaneSettingsRepository(path);
            return result;
        }

        private static Assembly ResolveAssemblyOnCurrentDomain(object sender, ResolveEventArgs args)
        {
            var requestedAssembly = new AssemblyName(args.Name);
            var assembly = default(Assembly);

            AppDomain.CurrentDomain.AssemblyResolve -= ResolveAssemblyOnCurrentDomain;

            try
            {
                assembly = Assembly.Load(requestedAssembly.Name);
            }
            catch
            { }

            AppDomain.CurrentDomain.AssemblyResolve += ResolveAssemblyOnCurrentDomain;

            return assembly;
        }

        #endregion Private Methods
    }
}

#pragma warning restore CA1031 // Do not catch general exception types