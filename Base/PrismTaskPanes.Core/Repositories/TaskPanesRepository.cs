#pragma warning disable CA1031 // Do not catch general exception types

using NetOffice.OfficeApi;
using PrismTaskPanes.Extensions;
using PrismTaskPanes.Settings;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;

namespace PrismTaskPanes.Factories
{
    [ComVisible(false)]
    public class TaskPanesRepository
        : IDisposable
    {
        #region Private Fields

        private readonly TaskPaneSettingsRepository configurationsRepository;
        private readonly Func<string> documentHashGetter;
        private readonly Dictionary<string, CustomTaskPane> taskPanes = new Dictionary<string, CustomTaskPane>();
        private readonly TaskPanesFactory taskPanesFactory;

        private bool disposed;

        #endregion Private Fields

        #region Public Constructors

        public TaskPanesRepository(int key, object scope, TaskPanesFactory taskPanesFactory,
            TaskPaneSettingsRepository configurationsRepository, Func<string> documentHashGetter)
        {
            Key = key;
            Scope = scope;

            this.taskPanesFactory = taskPanesFactory;
            this.configurationsRepository = configurationsRepository;
            this.documentHashGetter = documentHashGetter;
        }

        #endregion Public Constructors

        #region Public Properties

        public int Key { get; }

        public object Scope { get; }

        #endregion Public Properties

        #region Public Methods

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        public bool Exists(string receiverHash)
        {
            var taskPane = GetExistingTaskPane(receiverHash);

            if (taskPane == default)
                taskPane = GetNewTaskPane(receiverHash);

            return taskPane != default;
        }

        public void Initialise()
        {
            var visibleTaskPanes = configurationsRepository
                .Get(documentHashGetter.Invoke())
                .Where(a => a.Visible)
                .Where(a => !a.InvisibleAtStart).ToArray();

            foreach (var visibleTaskPane in visibleTaskPanes)
            {
                SetTaskPaneVisible(
                    receiverHash: visibleTaskPane.ReceiverHash,
                    isVisible: true);
            }
        }

        public bool IsVisible(string receiverHash)
        {
            var taskPane = GetExistingTaskPane(receiverHash);

            return taskPane?.Visible ?? false;
        }

        public void Save()
        {
            foreach (var taskPane in taskPanes)
            {
                SaveAttributes(
                    key: taskPane.Key,
                    taskPane: taskPane.Value);
            }
        }

        public void SetVisible(string receiverHash, bool isVisible)
        {
            SetTaskPaneVisible(
                receiverHash: receiverHash,
                isVisible: isVisible);
        }

        #endregion Public Methods

        #region Protected Methods

        protected virtual void Dispose(bool disposing)
        {
            if (!disposed)
            {
                if (disposing)
                {
                    foreach (var taskPane in taskPanes)
                    {
                        taskPane.Value.Visible = false;
                        taskPane.Value.Dispose();
                    }

                    taskPanes.Clear();
                }

                disposed = true;
            }
        }

        #endregion Protected Methods

        #region Private Methods

        private CustomTaskPane GetExistingTaskPane(string receiverHash)
        {
            var result = taskPanes.ContainsKey(receiverHash)
                ? taskPanes[receiverHash]
                : default;

            return result;
        }

        private CustomTaskPane GetNewTaskPane(string receiverHash)
        {
            var settings = configurationsRepository.Get(
                receiverHash: receiverHash,
                documentHash: documentHashGetter.Invoke());

            var result = taskPanesFactory.Get(settings);

            result.OnDispose += OnTaskPaneDispose;

            taskPanes.Add(
                key: receiverHash,
                value: result);

            return result;
        }

        private void OnTaskPaneDispose(NetOffice.OnDisposeEventArgs eventArgs)
        {
            var taskPane = taskPanes
                .Where(t => t.Value == eventArgs.Sender).SingleOrDefault();

            if (taskPane.Value != default)
            {
                SaveAttributes(
                    key: taskPane.Key,
                    taskPane: taskPane.Value);

                taskPanes.Remove(taskPane.Key);
            }
        }

        private void SaveAttributes(string key, CustomTaskPane taskPane)
        {
            // The workbook before close event is unsafe since it can be cancelled.
            // Therefore all taskpanes are removed at application disposal.
            // But it could be that the task pane does not exist anymore. This case must be catched here.

            try
            {
                configurationsRepository.Set(
                    receiverHash: key,
                    documentHash: documentHashGetter.Invoke(),
                    visible: taskPane.Visible,
                    width: taskPane.Width,
                    height: taskPane.Height,
                    dockPosition: taskPane.GetDockPosition());
            }
            catch (NetOffice.Exceptions.PropertyGetCOMException)
            { }

            configurationsRepository.Save();
        }

        private void SetTaskPaneVisible(string receiverHash, bool isVisible)
        {
            var taskPane = GetExistingTaskPane(receiverHash);

            if (taskPane == default && isVisible)
                taskPane = GetNewTaskPane(receiverHash);

            if (taskPane != default)
                taskPane.Visible = isVisible;
        }

        #endregion Private Methods
    }
}

#pragma warning restore CA1031 // Do not catch general exception types