#pragma warning disable CA1031 // Do not catch general exception types

using NetOffice.PowerPointApi;
using Prism.Ioc;
using System;
using System.Diagnostics;
using System.IO;
using System.Linq;

namespace PrismTaskPanes.DryIoc.PowerPoint
{
    internal class PowerPointApplication
        : DryIocApplication, IDisposable
    {
        #region Private Fields

        private const string ProcessName = "POWERPNT";

        private readonly Application application;

        #endregion Private Fields

        #region Public Constructors

        public PowerPointApplication(object application, object ctpFactoryInst)
            : base(application, ctpFactoryInst)
        {
            this.application = application as Application;

            this.application.AfterNewPresentationEvent += OnAfterNewElement;
            this.application.AfterPresentationOpenEvent += OnAfterOpenElement;
            this.application.PresentationSaveEvent += OnAfterSaveElement; ;
            this.application.PresentationBeforeCloseEvent += OnBeforeCloseElement;

            this.application.OnDispose += OnApplicationDispose;
        }

        #endregion Public Constructors

        #region Protected Properties

        protected override object TaskPaneWindow => GetActiveWindow();

        protected override int? TaskPaneWindowKey => GetActiveWindowKey();

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
            var result = default(int?);

            var activeWindowName = GetActiveWindow()?.Caption;

            if (!string.IsNullOrWhiteSpace(activeWindowName))
            {
                result = Process.GetProcessesByName(ProcessName)
                    .Where(p => p.MainWindowTitle == $"{activeWindowName} - PowerPoint")
                    .SingleOrDefault()?.MainWindowHandle.ToInt32();
            }

            return result;
        }

        private void OnAfterNewElement(Presentation pres)
        {
            OpenScope();
        }

        private void OnAfterOpenElement(Presentation pres)
        {
            OpenScope();
        }

        private void OnAfterSaveElement(Presentation pres)
        {
            SaveScope();
        }

        private void OnApplicationDispose(NetOffice.OnDisposeEventArgs eventArgs)
        {
            base.OnApplicationDispose();
        }

        private void OnBeforeCloseElement(Presentation pres, ref bool cancel)
        {
            SaveScope();
        }

        #endregion Private Methods
    }
}

#pragma warning restore CA1031 // Do not catch general exception types