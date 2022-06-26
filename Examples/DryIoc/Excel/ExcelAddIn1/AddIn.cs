#pragma warning disable IDE0060 // Nicht verwendete Parameter entfernen
#pragma warning disable RCS1163 // Unused parameter.

using DryIoc;
using ExampleView;
using NetOffice.ExcelApi;
using NetOffice.ExcelApi.Tools;
using NetOffice.OfficeApi;
using NetOffice.Tools;
using Prism.Ioc;
using Prism.Modularity;
using PrismTaskPanes.Attributes;
using PrismTaskPanes.DryIoc;
using PrismTaskPanes.DryIoc.Excel;
using PrismTaskPanes.Enums;
using PrismTaskPanes.EventArgs;
using PrismTaskPanes.Extensions;
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

        #region Private Fields

        private TaskPanesProvider provider;

        #endregion Private Fields

        #region Public Constructors

        public AddIn()
        {
            OnStartupComplete += OnAddinStartupComplete;
        }

        #endregion Public Constructors

        #region Public Methods

        [ComRegisterFunction]
        public static void Register(Type type)
        {
            RegisterFunction(type);
            type.RegisterTaskPaneHost();
        }

        [ComUnregisterFunction]
        public static void Unregister(Type type)
        {
            UnregisterFunction(type);
            type.UnregisterTaskPaneHost();
        }

        public override void CTPFactoryAvailable(object CTPFactoryInst)
        {
            base.CTPFactoryAvailable(CTPFactoryInst);

            provider = new TaskPanesProvider(
                receiver: this,
                officeApplication: Application,
                ctpFactoryInst: CTPFactoryInst);

            provider.OnConfigureModuleCatalogEvent += OnConfigureModuleCatalog;
            provider.OnRegisterTypesEvent += OnRegisterTypes;

            provider.OnScopeOpenedEvent += OnScopeOpened;
        }

        public void InvalidateRibbonUI()
        {
            RibbonUI?.Invalidate();
        }

        public void TooglePaneVisibleButton_Click(IRibbonControl control, bool pressed)
        {
            provider.SetTaskPaneVisibility(
                id: "1",
                isVisible: pressed,
                checkApplication: true);
        }

        public void TooglePaneVisibleButton_Click2(IRibbonControl control, bool pressed)
        {
            provider.SetTaskPaneVisibility(
                id: "2",
                isVisible: pressed,
                checkApplication: true);
        }

        public bool TooglePaneVisibleButton_GetPressed(IRibbonControl control)
        {
            provider.InitilializeTaskPane(
                checkApplication: true);

            return provider.TaskPaneIsVisible(
                id: "1");
        }

        public bool TooglePaneVisibleButton_GetPressed2(IRibbonControl control)
        {
            provider.InitilializeTaskPane(checkApplication: true);

            return provider.TaskPaneIsVisible(
                id: "2");
        }

        #endregion Public Methods

        #region Private Methods

        private void OnAddinStartupComplete(ref Array custom)
        {
            Application.WorkbookBeforeCloseEvent += OnWorkbookBeforeClose;

            Console.WriteLine($"Addin started in Excel Version {Application.Version}");
        }

        private void OnConfigureModuleCatalog(object sender, ProviderEventArgs<IModuleCatalog> e)
        {
            e.Content.AddModule<ExampleModule>(nameof(ExcelAddIn1));
        }

        private void OnRegisterTypes(object sender, ProviderEventArgs<Prism.Ioc.IContainerRegistry> e)
        {
            e.Content.RegisterInstance<IExampleClass>(new ExampleClass());

            e.Content.RegisterForNavigation<ExampleView.Views.ViewAView>();
            e.Content.RegisterForNavigation<ExampleView.Views.ViewAView>(typeof(ExampleView.Views.ViewAView).FullName);
        }

        private void OnScopeOpened(object sender, ProviderEventArgs<DryIoc.IResolverContext> e)
        {
            var test1 = provider.Scope.Resolve<IExampleClass>();

            Console.WriteLine(test1.Message);

            var test2 = e.Content.Resolve<IExampleClass>();

            if (test1 != test2)
            {
                throw new ApplicationException($"{test1} and {test2} must be equal.");
            }
        }

        private void OnWorkbookBeforeClose(Workbook wb, ref bool cancel)
        {
            Console.WriteLine("Workbook will be closed.");
        }

        #endregion Private Methods

    }
}

#pragma warning restore RCS1163 // Unused parameter.
#pragma warning restore IDE0060 // Nicht verwendete Parameter entfernen