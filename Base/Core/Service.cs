using NetOffice.OfficeApi;
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
    public static class Service
    {
        #region Private Fields

        private const string SettingsFile = "PrismTaskPanes.xml";

        private static readonly HashSet<ITaskPanesReceiver> receivers = new();

        #endregion Private Fields

        #region Public Events

        public static event EventHandler<TaskPaneEventArgs> OnTaskPaneChangedEvent;

        #endregion Public Events

        #region Public Properties

        public static TaskPaneSettingsRepository ConfigurationsRepository { get; } = GetConfigurationsRepository();

        #endregion Public Properties

        #region Public Methods

        public static void AddReceiver(ITaskPanesReceiver receiver)
        {
            var attributes = receiver
                .GetAttributes().ToArray();

            ConfigurationsRepository.AddAttributes(attributes);

            receivers.Add(receiver);
        }

        public static void InvalidateRibbonUI()
        {
            DoWithAllReceivers((r) => r.InvalidateRibbonUI());
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

        #endregion Private Methods
    }
}