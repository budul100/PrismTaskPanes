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
        : COMAddin, IPrismTaskPaneReceiver
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
            moduleCatalog.AddModule<TestViewLib.TestViewLibModule>(nameof(TestAddIn1));
        }

        public override void CTPFactoryAvailable(object CTPFactoryInst)
        {
            base.CTPFactoryAvailable(CTPFactoryInst);

            this.AddCTPReceiver(
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

        public void TooglePaneVisibleButton_Click(IRibbonControl control, bool pressed)
        {
            this.SetTaskPaneVisibility("1", pressed);
            RibbonUI.Invalidate();
        }

        public void TooglePaneVisibleButton_Click2(IRibbonControl control, bool pressed)
        {
            this.SetTaskPaneVisibility("2", pressed);
            RibbonUI.Invalidate();
        }

        public bool TooglePaneVisibleButton_GetPressed(IRibbonControl control)
        {
            return this.GetTaskPaneVisibility("1");
        }

        public bool TooglePaneVisibleButton_GetPressed2(IRibbonControl control)
        {
            return this.GetTaskPaneVisibility("2");
        }

        #endregion Public Methods

        #region Private Methods

        private void Addin_OnStartupComplete(ref Array custom)
        {
            PrismTaskPanesProvider.Initialize();

            Console.WriteLine($"Addin started in Excel Version {Application.Version}");
        }

        #endregion Private Methods
    }
}