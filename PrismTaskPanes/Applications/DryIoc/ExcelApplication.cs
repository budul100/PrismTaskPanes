using NetOffice.ExcelApi;
using Prism.Ioc;
using System;
using System.IO;

namespace PrismTaskPanes.Applications.DryIoc
{
    internal class ExcelApplication
        : OfficeApplication, IDisposable
    {
        #region Private Fields

        private readonly Application application;

        #endregion Private Fields

        #region Public Constructors

        public ExcelApplication(object application, object ctpFactoryInst)
            : base(application, ctpFactoryInst)
        {
            this.application = application as Application;

            this.application.NewWorkbookEvent += OnWorkbookNew;
            this.application.WorkbookOpenEvent += OnWorkbookOpen;
            this.application.WorkbookAfterSaveEvent += OnWorkbookSaveAfter;
            this.application.OnDispose += OnApplicationDispose;
        }

        #endregion Public Constructors

        #region Protected Properties

        protected override object TaskPaneWindow => application?.ActiveWindow;

        protected override int? TaskPaneWindowKey => application?.ActiveWindow?.Hwnd;

        #endregion Protected Properties

        #region Protected Methods

        protected override System.Windows.Window CreateShell()
        {
            return default;
        }

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
            var path = application.ActiveWorkbook?.FullName;
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

        private void OnApplicationDispose(NetOffice.OnDisposeEventArgs eventArgs)
        {
            CloseScope();
            Dispose();
        }

        private void OnWorkbookNew(Workbook wb)
        {
            OpenScope();
        }

        private void OnWorkbookOpen(Workbook wb)
        {
            OpenScope();
        }

        private void OnWorkbookSaveAfter(Workbook wb, bool success)
        {
            SaveScope();
        }

        #endregion Private Methods
    }
}