#pragma warning disable IDE0060 // Nicht verwendete Parameter entfernen

using DryIoc;
using NetOffice.ExcelApi;
using NetOffice.ExcelApi.Tools;
using NetOffice.OfficeApi;
using NetOffice.Tools;
using Prism.Ioc;
using Prism.Modularity;
using PrismTaskPanes.Attributes;
using PrismTaskPanes.DryIoc;
using PrismTaskPanes.DryIoc.EventArgs;
using PrismTaskPanes.Enums;
using PrismTaskPanes.Interfaces;
using System;
using System.Runtime.InteropServices;
using TestCommon;

namespace ExcelAddIn1
{
    [COMAddin("PrismTaskPanes Excel Example-AddIn 1", "This is an ExampleAddIn1 description.", LoadBehavior.LoadAtStartup),
        ProgId("PrismTaskPanes.ExcelAddIn1"),
        Guid("43BFA36D-D6A0-4558-8079-C07919C3CA73"),
        RegistryLocation(RegistrySaveLocation.LocalMachine),
        Codebase,
        ComVisible(true)]
    [CustomUI("RibbonUI.xml", true)]
    [PrismTaskPane("1", "ExampleAddin 1 A", typeof(ExampleView.Views.ViewAView), "ExampleRegion", invisibleAtStart: true, ScrollBarVertical = ScrollVisibility.Disabled)]
    [PrismTaskPane("2", "ExampleAddin 1 B", typeof(ExampleView.Views.ViewAView), "ExampleRegion", navigationValue: "abc")]
    public class AddIn
        : COMAddin, ITaskPanesReceiver
    {
        #region Public Constructors

        public AddIn()
        {
            OnStartupComplete += OnAddinStartupComplete;

            ExcelProvider.OnProviderReadyEvent += OnProviderReady;
            ExcelProvider.OnScopeOpenedEvent += OnScopeOpened;
            ExcelProvider.OnScopeInitializedEvent += OnScopeInitialized;
        }

        #endregion Public Constructors

        #region Public Methods

        [ComRegisterFunction]
        public static void Register(Type type)
        {
            RegisterFunction(type);

            ExcelProvider.RegisterProvider<AddIn>();
        }

        [ComUnregisterFunction]
        public static void Unregister(Type type)
        {
            UnregisterFunction(type);

            ExcelProvider.UnregisterProvider<AddIn>();
        }

        public void ConfigureModuleCatalog(IModuleCatalog moduleCatalog)
        {
            moduleCatalog.AddModule<ExampleView.ExampleModule>(nameof(ExcelAddIn1));
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
            var test = new ExampleClass();

            containerRegistry.RegisterInstance<IExampleClass>(test);
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

        private void OnAddinStartupComplete(ref Array custom)
        {
            Application.WorkbookBeforeCloseEvent += OnWorkbookBeforeClose;

            Console.WriteLine($"Addin started in Excel Version {Application.Version}");
        }

        private void OnProviderReady(object sender, EventArgs e)
        {
            var test = ExcelProvider.Container.Resolve<IExampleClass>();

            Console.WriteLine(test.Message);
        }

        private void OnScopeInitialized(object sender, ExcelEventArgs e)
        {
            if (e.Container != ExcelProvider.Container)
            {
                throw new ApplicationException($"{ExcelProvider.Container} and {e.Container} must be equal.");
            }

            var test1 = ExcelProvider.Container.Resolve<IExampleClass>();
            var test2 = e.Container.Resolve<IExampleClass>();

            if (test1 != test2)
            {
                throw new ApplicationException($"{test1} and {test2} must be equal.");
            }
        }

        private void OnScopeOpened(object sender, ExcelEventArgs e)
        {
            var test1 = ExcelProvider.Container.Resolve<IExampleClass>();
            var test2 = e.Container.Resolve<IExampleClass>();
        }

        private void OnWorkbookBeforeClose(Workbook wb, ref bool cancel)
        {
            Console.WriteLine($"Workbook will be closed.");
        }

        #endregion Private Methods
    }
}

#pragma warning restore IDE0060 // Nicht verwendete Parameter entfernen