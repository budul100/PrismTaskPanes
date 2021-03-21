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

            this.application.NewWorkbookEvent += OnElementNewAfter;
            this.application.WorkbookOpenEvent += OnElementOpenAfter;
            this.application.WorkbookAfterSaveEvent += OnElementSaveAfter;
            this.application.WorkbookBeforeCloseEvent += OnElementCloseBefore;

            this.application.OnDispose += OnApplicationDispose;
        }

        #endregion Public Constructors

        #region Protected Properties

        protected override object TaskPaneWindow => GetActiveWindow();

        protected override int? TaskPaneWindowKey => GetActiveWindow()?.Hwnd;

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

        private Workbook GetActiveElement()
        {
            var result = default(Workbook);

            try
            {
                result = application?.ActiveWorkbook;
            }
            catch (NetOffice.Exceptions.PropertyGetCOMException)
            { }

            return result;
        }

        private Window GetActiveWindow()
        {
            var result = default(Window);

            try
            {
                result = application?.ActiveWindow;
            }
            catch (NetOffice.Exceptions.PropertyGetCOMException)
            { }

            return result;
        }

        private void OnElementCloseBefore(Workbook wb, ref bool cancel)
        {
            SaveScope();
        }

        private void OnElementNewAfter(Workbook wb)
        {
            OpenScope();
        }

        private void OnElementOpenAfter(Workbook wb)
        {
            OpenScope();
        }

        private void OnElementSaveAfter(Workbook wb, bool success)
        {
            SaveScope();
        }

        #endregion Private Methods
    }
}