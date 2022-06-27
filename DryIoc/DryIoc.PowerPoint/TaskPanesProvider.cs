using NetOffice.PowerPointApi;
using PrismTaskPanes.Extensions;
using PrismTaskPanes.Interfaces;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;

namespace PrismTaskPanes.DryIoc.PowerPoint
{
    public class TaskPanesProvider
        : BaseProvider
    {
        #region Private Fields

        private const string ProcessName = "POWERPNT";

        private readonly IDictionary<string, int> handles = new Dictionary<string, int>();
        private readonly NetOffice.PowerPointApi.Application officeApplication;

        #endregion Private Fields

        #region Public Constructors

        public TaskPanesProvider(ITaskPanesReceiver receiver, NetOffice.PowerPointApi.Application officeApplication,
            object ctpFactoryInst, bool suppressInitializationAtStart = false, bool showErrorIfAlreadyLoaded = false)
            : base(receiver: receiver, officeApplication: officeApplication, ctpFactoryInst: ctpFactoryInst,
                  suppressInitializationAtStart: suppressInitializationAtStart, showErrorIfAlreadyLoaded: showErrorIfAlreadyLoaded)
        {
            if (officeApplication is null)
            {
                throw new ArgumentNullException(nameof(officeApplication));
            }

            if (ctpFactoryInst is null)
            {
                throw new ArgumentNullException(nameof(ctpFactoryInst));
            }

            this.officeApplication = officeApplication;

            this.officeApplication.AfterNewPresentationEvent += OnAfterNewElement;
            this.officeApplication.AfterPresentationOpenEvent += OnAfterOpenElement;
            this.officeApplication.PresentationSaveEvent += OnAfterSaveElement;
            this.officeApplication.PresentationCloseFinalEvent += OnAfterCloseEvent;
            this.officeApplication.OnDispose += OnApplicationDispose;

            TaskPaneWindowGetter = GetActiveWindow;
            TaskPaneWindowKeyGetter = GetActiveWindowKey;
            TaskPaneIdentifierGetter = GetTaskPaneIdentifier;
        }

        #endregion Public Constructors

        #region Protected Methods

        protected override void Dispose(bool disposing)
        {
            if (!IsDisposed)
            {
                if (disposing)
                {
                    officeApplication?.DisposeChildInstances();
                }

                base.Dispose(disposing);
            }
        }

        #endregion Protected Methods

        #region Private Methods

        private Presentation GetActiveElement()
        {
            var result = default(Presentation);

            try
            {
                result = officeApplication?.ActivePresentation;
            }
            catch (NetOffice.Exceptions.PropertyGetCOMException)
            { }

            return result;
        }

        private DocumentWindow GetActiveWindow()
        {
            var result = default(DocumentWindow);

            try
            {
                result = officeApplication?.ActiveWindow;
            }
            catch (NetOffice.Exceptions.PropertyGetCOMException)
            { }

            return result;
        }

        private int? GetActiveWindowKey()
        {
            var result = default(int);

            var processes = Process.GetProcessesByName(ProcessName).ToArray();

            if (processes.Length > 0)
            {
                var activeWindowName = GetActiveWindow()?.Caption;
                var key = $"{activeWindowName} - PowerPoint";

                var process = processes
                    .SingleOrDefault(p => p.MainWindowTitle == key);

                if (process?.MainWindowTitle == key)
                {
                    if (handles.ContainsKey(key))
                    {
                        handles[key] = process.MainWindowHandle.ToInt32();
                    }
                    else
                    {
                        handles.Add(
                            key: key,
                            value: process.MainWindowHandle.ToInt32());
                    }
                }

                if (handles.ContainsKey(key))
                {
                    result = handles[key];
                }
            }

            return result;
        }

        private string GetTaskPaneIdentifier()
        {
            var path = GetActiveElement()?.FullName;
            var isPath = !string.IsNullOrWhiteSpace(Path.GetDirectoryName(path));

            return isPath ? path.GetHashString() : default;
        }

        private void OnAfterCloseEvent(Presentation pres)
        {
            SaveScope();

            Service.InvalidateRibbonUI();
        }

        private void OnAfterNewElement(Presentation pres)
        {
            OpenScope();

            Service.InvalidateRibbonUI();
        }

        private void OnAfterOpenElement(Presentation pres)
        {
            OpenScope();

            Service.InvalidateRibbonUI();
        }

        private void OnAfterSaveElement(Presentation pres)
        {
            SaveScope();
        }

        #endregion Private Methods
    }
}