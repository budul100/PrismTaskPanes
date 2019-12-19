﻿using NetOffice.OfficeApi;
using System;
using System.Collections.Generic;
using System.Linq;

namespace PrismTaskPanes.TaskPanes
{
    internal class TaskPanesRepository :
        IDisposable
    {
        #region Private Fields

        private readonly Func<int> documentHashGetter;
        private readonly SettingsRepository settingsRepository;
        private readonly Dictionary<int, CustomTaskPane> taskPanes = new Dictionary<int, CustomTaskPane>();
        private readonly TaskPanesFactory taskPanesFactory;

        private bool disposed = false;

        #endregion Private Fields

        #region Public Constructors

        public TaskPanesRepository(int key, TaskPanesFactory taskPanesFactory, SettingsRepository settingsRepository,
            Func<int> documentHashGetter)
        {
            Key = key;

            this.taskPanesFactory = taskPanesFactory;
            this.settingsRepository = settingsRepository;
            this.documentHashGetter = documentHashGetter;

            OpenTaskPanes();
        }

        #endregion Public Constructors

        #region Public Properties

        public int Key { get; private set; }

        #endregion Public Properties

        #region Public Methods

        public void Dispose()
        {
            Dispose(true);

            GC.SuppressFinalize(this);
        }

        public bool GetVisibility(int hash)
        {
            var taskPane = GetExistingTaskPane(hash);

            return taskPane?.Visible ?? false;
        }

        public void Save()
        {
            SaveAttributes();
        }

        public void SetVisibility(int hash, bool isVisible)
        {
            SetTaskPaneVisible(
                hash: hash,
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

        private CustomTaskPane GetExistingTaskPane(int hash)
        {
            var result = taskPanes.ContainsKey(hash)
                ? taskPanes[hash]
                : default;

            return result;
        }

        private CustomTaskPane GetNewTaskPane(int hash)
        {
            var settings = settingsRepository.Get(
                attributeHash: hash,
                documentHash: documentHashGetter.Invoke());

            var result = taskPanesFactory.Get(settings);

            taskPanes.Add(
                key: hash,
                value: result);

            return result;
        }

        private void OpenTaskPanes()
        {
            var visibleTaskPanes = settingsRepository
                .Get(documentHashGetter.Invoke())
                .Where(a => a.Visible)
                .Where(a => !a.InvisibleAtStart).ToArray();

            foreach (var visibleTaskPane in visibleTaskPanes)
            {
                SetTaskPaneVisible(
                    hash: visibleTaskPane.AttributeHash,
                    isVisible: true);
            }
        }

        private void SaveAttributes()
        {
            foreach (var taskPane in taskPanes)
            {
                settingsRepository.Set(
                    attributeHash: taskPane.Key,
                    documentHash: documentHashGetter.Invoke(),
                    visible: taskPane.Value.Visible,
                    width: taskPane.Value.Width,
                    height: taskPane.Value.Height,
                    dockPosition: taskPane.Value.DockPosition);
            }

            settingsRepository.Save();
        }

        private void SetTaskPaneVisible(int hash, bool isVisible)
        {
            var taskPane = GetExistingTaskPane(hash);

            if (taskPane == null && isVisible)
                taskPane = GetNewTaskPane(hash);

            if (taskPane != null)
                taskPane.Visible = isVisible;
        }

        #endregion Private Methods
    }
}