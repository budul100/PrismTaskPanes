using DryIoc;
using NetOffice.OfficeApi;
using NetOffice.PowerPointApi;
using NetOffice.PowerPointApi.Tools;
using NetOffice.Tools;
using Prism.DryIoc;
using Prism.Ioc;
using Prism.Modularity;
using PrismTaskPanes.Attributes;
using PrismTaskPanes.DryIoc;
using PrismTaskPanes.DryIoc.EventArgs;
using PrismTaskPanes.Interfaces;
using System;
using System.Runtime.InteropServices;
using TestCommon;

namespace PowerPointAddIn1
{
    [COMAddin("PrismTaskPanes.PowerPointAddIn1", "This is an PowerPointAddIn1 description.", LoadBehavior.LoadAtStartup),
        ProgId("PowerPointAddIn1.AddIn"),
        Guid("6EAFFBCE-AB05-4DC2-9709-B12595A192FF"),
        RegistryLocation(RegistrySaveLocation.CurrentUser)]
    [CustomUI("RibbonUI.xml", true)]
    [PrismTaskPane("1", "ExampleAddin 1 A", typeof(ExampleView.Views.ViewAView), "ExampleRegion", invisibleAtStart: true)]
    [PrismTaskPane("2", "ExampleAddin 1 B", typeof(ExampleView.Views.ViewAView), "ExampleRegion", navigationValue: "abc")]
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

        [ComRegisterFunction]
        public static void Register(Type type)
        {
            RegisterFunction(type);

            PowerPointProvider.RegisterProvider<AddIn>();
        }

        [ComUnregisterFunction]
        public static void Unregister(Type type)
        {
            UnregisterFunction(type);

            PowerPointProvider.UnregisterProvider<AddIn>();
        }

        public void ConfigureModuleCatalog(IModuleCatalog moduleCatalog)
        {
            moduleCatalog.AddModule<ExampleView.Module>(nameof(PowerPointAddIn1));
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

        public bool TooglePaneVisibleButton_GetEnabled(IRibbonControl control)
        {
            return this.TaskPaneExists("1");
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
            PowerPointProvider.OnScopeOpenedEvent += OnScopeOpened;
            PowerPointProvider.OnScopeInitializedEvent += OnScopeInitialized;

            Console.WriteLine($"Addin started in Excel Version {Application.Version}");
        }

        private void OnScopeInitialized(object sender, PowerPointEventArgs e)
        {
            var test1 = PowerPointProvider.Container.Resolve<IExampleClass>();
            var test2 = e.Container.Resolve<IExampleClass>();
        }

        private void OnScopeOpened(object sender, PowerPointEventArgs e)
        {
            var test1 = PowerPointProvider.Container.Resolve<IExampleClass>();
            var test2 = e.Container.Resolve<IExampleClass>();
        }

        #endregion Private Methods
    }
}