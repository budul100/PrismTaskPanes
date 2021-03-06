﻿#pragma warning disable CA1031 // Do not catch general exception types

using NetOffice.ExcelApi;
using Prism.Ioc;
using System;
using System.IO;

namespace PrismTaskPanes.DryIoc.Application
{
    internal class ExcelApplication
        : DryIocApplication, IDisposable
    {
        #region Private Fields

        private readonly NetOffice.ExcelApi.Application application;

        #endregion Private Fields

        #region Public Constructors

        public ExcelApplication(object application, object ctpFactoryInst, Type contentType)
            : base(application, ctpFactoryInst, contentType)
        {
            this.application = application as NetOffice.ExcelApi.Application;

            this.application.NewWorkbookEvent += OnAfterNewElement;
            this.application.WorkbookOpenEvent += OnAfterOpenElement;
            this.application.WorkbookAfterSaveEvent += OnAfterSaveElement;
            this.application.WorkbookBeforeCloseEvent += OnBeforeCloseElement;

            this.application.OnDispose += OnApplicationDispose;
        }

        #endregion Public Constructors

        #region Protected Properties

        protected override Func<object> TaskPaneWindowGetter => () => GetActiveWindow();

        protected override Func<int?> TaskPaneWindowKeyGetter => () => GetActiveWindow()?.Hwnd;

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

        private void OnAfterNewElement(Workbook wb)
        {
            OpenScope();

            BaseProvider.InvalidateRibbonUI();
        }

        private void OnAfterOpenElement(Workbook wb)
        {
            OpenScope();

            BaseProvider.InvalidateRibbonUI();
        }

        private void OnAfterSaveElement(Workbook wb, bool success)
        {
            SaveScope();
        }

        private void OnApplicationDispose(NetOffice.OnDisposeEventArgs eventArgs)
        {
            base.OnApplicationDispose();
        }

        private void OnBeforeCloseElement(Workbook wb, ref bool cancel)
        {
            SaveScope();

            BaseProvider.InvalidateRibbonUI();
        }

        #endregion Private Methods
    }
}

#pragma warning restore CA1031 // Do not catch general exception types