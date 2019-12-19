using NetOffice.ExcelApi;
using PrismTaskPanes.TaskPanes;
using System.IO;

namespace PrismTaskPanes.AddIns.DryIOC
{
    internal class ExcelApplication
        : OfficeApplication
    {
        #region Private Fields

        private readonly Application application;

        #endregion Private Fields

        #region Public Constructors

        public ExcelApplication(object application, object ctpFactoryInst, SettingsRepository settingsRepository) :
            base(application, ctpFactoryInst, settingsRepository)
        {
            this.application = application as Application;

            this.application.WorkbookOpenEvent += OnWorkbookOpen;
            this.application.NewWorkbookEvent += OnWorkbookNew;
            this.application.WorkbookAfterSaveEvent += OnWorkbookSaveAfter;
            this.application.WorkbookBeforeCloseEvent += OnWorkbookBeforeClose; ;
        }

        #endregion Public Constructors

        #region Protected Properties

        protected override object TaskPaneWindow => application?.ActiveWindow;

        protected override int? TaskPaneWindowKey => application?.ActiveWindow?.Hwnd;

        #endregion Protected Properties

        #region Public Methods

        public override void Dispose()
        {
            application.WorkbookOpenEvent -= OnWorkbookOpen;
            application.NewWorkbookEvent -= OnWorkbookNew;
            application.WorkbookBeforeCloseEvent -= OnWorkbookBeforeClose;
            application.WorkbookAfterSaveEvent -= OnWorkbookSaveAfter;

            base.Dispose();
        }

        #endregion Public Methods

        #region Protected Methods

        protected override string GetTaskPaneIdentifier()
        {
            var path = application.ActiveWorkbook?.FullName ?? null;
            var isPath = !string.IsNullOrWhiteSpace(Path.GetDirectoryName(path));

            return isPath
                ? path
                : default;
        }

        #endregion Protected Methods

        #region Private Methods

        private void OnWorkbookBeforeClose(Workbook wb, ref bool cancel)
        {
            if (!cancel) CloseWindow();
        }

        private void OnWorkbookNew(Workbook wb)
        {
            isActivated = true;
            GetRepository();
        }

        private void OnWorkbookOpen(Workbook wb)
        {
            isActivated = true;
            GetRepository();
        }

        private void OnWorkbookSaveAfter(Workbook wb, bool success)
        {
            GetRepository().Save();
        }

        #endregion Private Methods
    }
}