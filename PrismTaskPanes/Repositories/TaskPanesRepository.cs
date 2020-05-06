using NetOffice.OfficeApi;
using PrismTaskPanes.Settings;
using System;
using System.Collections.Generic;
using System.Linq;

namespace PrismTaskPanes.Factories
{
    internal class TaskPanesRepository :
        IDisposable
    {
        #region Private Fields

        private readonly TaskPaneSettingsRepository configurationsRepository;
        private readonly Func<string> documentHashGetter;
        private readonly Dictionary<string, CustomTaskPane> taskPanes = new Dictionary<string, CustomTaskPane>();
        private readonly TaskPanesFactory taskPanesFactory;

        private bool disposed = false;

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
            SaveAttributes();
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

            taskPanes.Add(
                key: receiverHash,
                value: result);

            return result;
        }

        private void SaveAttributes()
        {
            foreach (var taskPane in taskPanes)
            {
                configurationsRepository.Set(
                    receiverHash: taskPane.Key,
                    documentHash: documentHashGetter.Invoke(),
                    visible: taskPane.Value.Visible,
                    width: taskPane.Value.Width,
                    height: taskPane.Value.Height,
                    dockPosition: taskPane.Value.DockPosition);
            }

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