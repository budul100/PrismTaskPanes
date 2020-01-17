using NetOffice.ExcelApi.Tools;
using NetOffice.OfficeApi;
using NetOffice.Tools;
using Prism.Ioc;
using Prism.Modularity;
using PrismTaskPanes;
using PrismTaskPanes.Attributes;
using PrismTaskPanes.Interfaces;
using System;
using System.Runtime.InteropServices;
using TestViewLib.Views;

namespace TestAddIn2
{
    [COMAddin("PrismTaskPanes.TestAddIn2", "PrismTaskPanes.TestAddIn2", LoadBehavior.LoadAtStartup),
        ProgId("PrismTaskPanes.TestAddIn2"),
        Guid("4B20A71A-E9C3-47AD-8E3A-61EB50B51AAE"),
        ComVisible(true),
        Codebase]
    [CustomUI("RibbonUI.xml", true),
        RegistryLocation(RegistrySaveLocation.LocalMachine)]
    [PrismTaskPane("1", "TestAddin 2 A", typeof(ViewA), "TestRegion", invisibleAtStart: true)]
    [PrismTaskPane("2", "TestAddin 2 B", typeof(ViewA), "TestRegion",
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
            moduleCatalog.AddModule<TestViewLib.Module>(nameof(TestAddIn2));
        }

        public override void CTPFactoryAvailable(object CTPFactoryInst)
        {
            base.CTPFactoryAvailable(CTPFactoryInst);

            this.InitializeTaskPanesProvider(
                application: Application,
                ctpFactoryInst: CTPFactoryInst);
        }

        public void InvalidateRibbonUI()
        {
            RibbonUI.Invalidate();
        }

        public void RegisterTypes(IContainerRegistry builder)
        {
        }

        public void RegisterTypes(IContainerProvider containerProvider)
        { }

        public void TooglePaneVisibleButton_Click(IRibbonControl control, bool pressed)
        {
            "1".SetTaskPaneVisible(pressed);
        }

        public void TooglePaneVisibleButton_Click2(IRibbonControl control, bool pressed)
        {
            "2".SetTaskPaneVisible(pressed);
        }

        public bool TooglePaneVisibleButton_GetPressed(IRibbonControl control)
        {
            return "1".TaskPaneVisible();
        }

        public bool TooglePaneVisibleButton_GetPressed2(IRibbonControl control)
        {
            return "2".TaskPaneVisible();
        }

        #endregion Public Methods

        #region Private Methods

        private void Addin_OnStartupComplete(ref Array custom)
        {
            Console.WriteLine($"Addin started in Excel Version {Application.Version}");
        }

        #endregion Private Methods
    }
}