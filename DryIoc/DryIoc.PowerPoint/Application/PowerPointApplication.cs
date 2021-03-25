#pragma warning disable CA1031 // Do not catch general exception types

using NetOffice.PowerPointApi;
using Prism.Ioc;
using System;
using System.Diagnostics;
using System.IO;
using System.Linq;

namespace PrismTaskPanes.DryIoc.Application
{
    internal class PowerPointApplication
        : DryIocApplication, IDisposable
    {
        #region Private Fields

        private const string ProcessName = "POWERPNT";

        private readonly NetOffice.PowerPointApi.Application application;

        #endregion Private Fields

        #region Public Constructors

        public PowerPointApplication(object application, object ctpFactoryInst)
            : base(application, ctpFactoryInst)
        {
            this.application = application as NetOffice.PowerPointApi.Application;

            this.application.AfterNewPresentationEvent += OnAfterNewElement;
            this.application.AfterPresentationOpenEvent += OnAfterOpenElement;
            this.application.PresentationSaveEvent += OnAfterSaveElement; ;
            this.application.PresentationCloseEvent += OnAfterCloseEvent;

            this.application.OnDispose += OnApplicationDispose;
        }

        #endregion Public Constructors

        #region Protected Properties

        protected override Func<object> TaskPaneWindowGetter => () => GetActiveWindow();

        protected override Func<int?> TaskPaneWindowKeyGetter => () => GetActiveWindowKey();

        #endregion Protected Properties

        #region Protected Methods

        protected override void Dispose(bool disposing)
        {
            if (!isDisposed)
            {
                if (disposing)
                {
                    application?.DisposeChildInstances();
                }

                base.Dispose(disposing);
            }
        }

        protected override string GetTaskPaneIdentifier()
        {
            var path = GetActiveElement()?.FullName;
            var isPath = !string.IsNullOrWhiteSpace(Path.GetDirectoryName(path));

            return isPath
                ? path
                : default;
        }

        protected override void RegisterTypes(IContainerRegistry containerRegistry)
        {
            base.RegisterTypes(containerRegistry);

            containerRegistry.RegisterInstance(application);
        }

        #endregion Protected Methods

        #region Private Methods

        private Presentation GetActiveElement()
        {
            var result = default(Presentation);

            try
            {
                result = application?.ActivePresentation;
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
                result = application?.ActiveWindow;
            }
            catch (NetOffice.Exceptions.PropertyGetCOMException)
            { }

            return result;
        }

        private int? GetActiveWindowKey()
        {
            var result = default(Process);

            var processes = Process.GetProcessesByName(ProcessName).ToArray();

            if (processes.Count() == 1)
            {
                result = processes.SingleOrDefault();
            }
            else if (processes.Any())
            {
                var activeWindowName = GetActiveWindow()?.Caption;

                result = processes
                    .Where(p => p.MainWindowTitle == $"{activeWindowName} - PowerPoint").SingleOrDefault();

                if (result == default)
                {
                    result = processes
                        .Where(p => string.IsNullOrWhiteSpace(p.MainWindowTitle)).SingleOrDefault();
                }
            }

            return result?.MainWindowHandle.ToInt32();
        }

        private void OnAfterCloseEvent(Presentation pres)
        {
            SaveScope();

            BaseProvider.InvalidateRibbonUI();
        }

        private void OnAfterNewElement(Presentation pres)
        {
            OpenScope();

            BaseProvider.InvalidateRibbonUI();
        }

        private void OnAfterOpenElement(Presentation pres)
        {
            OpenScope();

            BaseProvider.InvalidateRibbonUI();
        }

        private void OnAfterSaveElement(Presentation pres)
        {
            SaveScope();
        }

        private void OnApplicationDispose(NetOffice.OnDisposeEventArgs eventArgs)
        {
            base.OnApplicationDispose();
        }

        #endregion Private Methods
    }
}

#pragma warning restore CA1031 // Do not catch general exception types