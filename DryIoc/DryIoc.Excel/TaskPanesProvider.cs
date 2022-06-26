using NetOffice.ExcelApi;
using PrismTaskPanes.Extensions;
using PrismTaskPanes.Interfaces;
using System;
using System.IO;

namespace PrismTaskPanes.DryIoc.Excel
{
    public class TaskPanesProvider
        : BaseProvider
    {
        #region Private Fields

        private readonly NetOffice.ExcelApi.Application officeApplication;

        #endregion Private Fields

        #region Public Constructors

        public TaskPanesProvider(ITaskPanesReceiver receiver, NetOffice.ExcelApi.Application officeApplication,
            object ctpFactoryInst)
            : base(receiver, officeApplication, ctpFactoryInst)
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

            this.officeApplication.NewWorkbookEvent += OnAfterNewElement;
            this.officeApplication.WorkbookOpenEvent += OnAfterOpenElement;
            this.officeApplication.WorkbookAfterSaveEvent += OnAfterSaveElement;
            this.officeApplication.WorkbookBeforeCloseEvent += OnBeforeCloseElement;
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

        private Workbook GetActiveElement()
        {
            var result = default(Workbook);

            try
            {
                result = officeApplication?.ActiveWorkbook;
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
                result = officeApplication?.ActiveWindow;
            }
            catch (NetOffice.Exceptions.PropertyGetCOMException)
            { }

            return result;
        }

        private int? GetActiveWindowKey()
        {
            return GetActiveWindow()?.Hwnd;
        }

        private string GetTaskPaneIdentifier()
        {
            var path = GetActiveElement()?.FullName;
            var isPath = !string.IsNullOrWhiteSpace(Path.GetDirectoryName(path));

            return isPath ? path.GetHashString() : default;
        }

        private void OnAfterNewElement(Workbook wb)
        {
            OpenScope();

            Service.InvalidateRibbonUI();
        }

        private void OnAfterOpenElement(Workbook wb)
        {
            OpenScope();

            Service.InvalidateRibbonUI();
        }

        private void OnAfterSaveElement(Workbook wb, bool success)
        {
            SaveScope();
        }

        private void OnBeforeCloseElement(Workbook wb, ref bool cancel)
        {
            SaveScope();

            Service.InvalidateRibbonUI();
        }

        #endregion Private Methods
    }
}