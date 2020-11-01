﻿using DryIoc;
using NetOffice.ExcelApi;
using NetOffice.ExcelApi.Tools;
using NetOffice.OfficeApi;
using NetOffice.Tools;
using Prism.DryIoc;
using Prism.Ioc;
using Prism.Modularity;
using PrismTaskPanes;
using PrismTaskPanes.Attributes;
using PrismTaskPanes.Interfaces;
using System;
using System.Runtime.InteropServices;
using TestCommon;

namespace ExampleAddIn1
{
    [COMAddin("PrismTaskPanes.ExampleAddIn1", "This is an ExampleAddIn1 description.", LoadBehavior.LoadAtStartup),
        ProgId("ExampleAddIn1.AddIn"),
        Guid("9F60AAAD-D3A1-40CF-9089-D46C000C75E0"),
        ComVisible(true),
        Codebase]
    [CustomUI("RibbonUI.xml", true),
        RegistryLocation(RegistrySaveLocation.LocalMachine)]
    [PrismTaskPane("1", "ExampleAddin 1 A", typeof(ExampleView.Views.ViewAView), "ExampleRegion1", invisibleAtStart: true)]
    [PrismTaskPane("2", "ExampleAddin 1 B", typeof(ExampleView.Views.ViewAView), "ExampleRegion")]
    public class AddIn
        : COMAddin, ITaskPanesReceiver
    {
        #region Public Constructors

        public AddIn()
        {
            OnStartupComplete += Addin_OnStartupComplete;
        }

        #endregion Public Constructors

        #region Public Methods

        public void ConfigureModuleCatalog(IModuleCatalog moduleCatalog)
        {
            moduleCatalog.AddModule<ExampleView.Module>(nameof(ExampleAddIn1));
        }

        public override void CTPFactoryAvailable(object CTPFactoryInst)
        {
            base.CTPFactoryAvailable(CTPFactoryInst);

            this.InitializeProvider(
                application: Application,
                ctpFactoryInst: CTPFactoryInst);
        }

        public void InvalidateRibbonUI()
        {
            RibbonUI?.Invalidate();
        }

        public void RegisterTypes(IContainerRegistry containerRegistry)
        {
            containerRegistry.GetContainer().Register<IExampleClass, ExampleClass>();
        }

        public void TooglePaneVisibleButton_Click(IRibbonControl control, bool pressed)
        {
            this.SetTaskPaneVisible(
                id: "1",
                isVisible: pressed);
        }

        public void TooglePaneVisibleButton_Click2(IRibbonControl control, bool pressed)
        {
            this.SetTaskPaneVisible(
                id: "2",
                isVisible: pressed);
        }

        public bool TooglePaneVisibleButton_GetPressed(IRibbonControl control)
        {
            return this.TaskPaneVisible("1");
        }

        public bool TooglePaneVisibleButton_GetPressed2(IRibbonControl control)
        {
            return this.TaskPaneVisible("2");
        }

        #endregion Public Methods

        #region Private Methods

        private void Addin_OnStartupComplete(ref Array custom)
        {
            Application.WorkbookBeforeCloseEvent += OnWorkbookBeforeClose;

            DryIocProvider.OnScopeOpenedEvent += OnScopeOpened;
            DryIocProvider.OnScopeInitialized += OnScopeInitialized;

            Console.WriteLine($"Addin started in Excel Version {Application.Version}");
        }

        private void OnScopeInitialized(object sender, PrismTaskPanes.Events.DryIocEventArgs e)
        {
            var test1 = DryIocProvider.Container.Resolve<IExampleClass>();
            var test2 = e.Container.Resolve<IExampleClass>();
        }

        private void OnScopeOpened(object sender, PrismTaskPanes.Events.DryIocEventArgs e)
        {
            var test1 = DryIocProvider.Container.Resolve<IExampleClass>();
            var test2 = e.Container.Resolve<IExampleClass>();
        }

        private void OnWorkbookBeforeClose(Workbook wb, ref bool cancel)
        {
            Console.WriteLine($"Workbook will be closed.");
        }

        #endregion Private Methods
    }
}