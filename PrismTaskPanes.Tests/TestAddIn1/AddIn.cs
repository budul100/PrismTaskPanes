using DryIoc;
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
using TestViewLib.Views;

namespace TestAddIn1
{
    [COMAddin("PrismTaskPanes.TestAddIn1", "PrismTaskPanes.TestAddIn1", LoadBehavior.LoadAtStartup),
        ProgId("PrismTaskPanes.TestAddIn1"),
        Guid("37434A4F-3ADE-4FF0-BD1B-92A479FFA4AE"),
        ComVisible(true),
        Codebase]
    [CustomUI("RibbonUI.xml", true),
        RegistryLocation(RegistrySaveLocation.LocalMachine)]
    [PrismTaskPane("1", "TestAddin 1 A", typeof(ViewA), "TestRegion", invisibleAtStart: true)]
    [PrismTaskPane("2", "TestAddin 1 B", typeof(ViewA), "TestRegion",
        navigationKey: "x", navigationValue: "")]
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
            moduleCatalog.AddModule<TestViewLib.Module>(nameof(TestAddIn1));
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
            containerRegistry.GetContainer().Register<ITestClass, TestClass>();
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
            var test = e.Container.Resolve<ITestClass>();
        }

        private void OnScopeOpened(object sender, PrismTaskPanes.Events.DryIocEventArgs e)
        {
            var test = e.Container.Resolve<ITestClass>();
        }

        private void OnWorkbookBeforeClose(Workbook wb, ref bool cancel)
        {
            Console.WriteLine($"Workbook will be closed.");
        }

        #endregion Private Methods
    }
}