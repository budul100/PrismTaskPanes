#pragma warning disable IDE0060 // Nicht verwendete Parameter entfernen
#pragma warning disable RCS1163 // Unused parameter.

using DryIoc;
using ExampleCommon;
using NetOffice.OfficeApi;
using NetOffice.PowerPointApi.Tools;
using NetOffice.Tools;
using Prism.Ioc;
using Prism.Modularity;
using PrismTaskPanes.Attributes;
using PrismTaskPanes.DryIoc.PowerPoint;
using PrismTaskPanes.EventArgs;
using PrismTaskPanes.Extensions;
using PrismTaskPanes.Interfaces;
using System;
using System.Runtime.InteropServices;
using TestCommon;

namespace PowerPointAddIn1
{
    [COMAddin("PrismTaskPanes.PowerPointAddIn1", "This is an PowerPointAddIn1 description.", LoadBehavior.LoadAtStartup),
        ProgId("PowerPointAddIn1.AddIn"),
        Guid("6EAFFBCE-AB05-4DC2-9709-B12595A192FF"),
        RegistryLocation(RegistrySaveLocation.LocalMachine),
        Codebase,
        ComVisible(true)]
    [CustomUI("RibbonUI.xml", true)]
    [PrismTaskPane("1", "ExampleAddin 1 A", typeof(ExampleView.Views.ViewAView), "ExampleRegion", invisibleAtStart: true)]
    [PrismTaskPane("2", "ExampleAddin 1 B", typeof(ExampleView.Views.ViewAView), "ExampleRegion", navigationValue: "abc")]
    public class AddIn
        : COMAddin, ITaskPanesReceiver, IApplication
    {
        #region Private Fields

        private TaskPanesProvider provider;

        #endregion Private Fields

        #region Public Constructors

        public AddIn()
        {
            //OnStartupComplete += Addin_OnStartupComplete;
        }

        #endregion Public Constructors

        #region Public Methods

        [ComRegisterFunction]
        public static void Register(Type type)
        {
            RegisterFunction(type);
            //type.RegisterTaskPaneHost();
        }

        [ComUnregisterFunction]
        public static void Unregister(Type type)
        {
            UnregisterFunction(type);
            type.UnregisterTaskPaneHost();
        }

        public void CallDispatcher(System.Action callback)
        {
            TaskPanesProvider.CallDispatcher(callback);
        }

        public override void CTPFactoryAvailable(object CTPFactoryInst)
        {
            provider = new TaskPanesProvider(
                receiver: this,
                officeApplication: Application,
                ctpFactoryInst: CTPFactoryInst,
                showErrorIfAlreadyLoaded: true);

            provider.OnConfigureModuleCatalogEvent += OnConfigureModuleCatalog;
            provider.OnRegisterTypesEvent += OnRegisterTypes;
        }

        public void InvalidateRibbonUI()
        {
            RibbonUI?.Invalidate();
        }

        public void TooglePaneVisibleButton_Click(IRibbonControl control, bool pressed)
        {
            provider.SetTaskPaneVisibility(
                id: "1",
                isVisible: pressed);
        }

        public void TooglePaneVisibleButton_Click2(IRibbonControl control, bool pressed)
        {
            provider.SetTaskPaneVisibility(
                id: "2",
                isVisible: pressed);
        }

        public bool TooglePaneVisibleButton_GetEnabled(IRibbonControl control)
        {
            return provider.TaskPaneExists("1");
        }

        public bool TooglePaneVisibleButton_GetPressed(IRibbonControl control)
        {
            return provider.TaskPaneIsVisible("1");
        }

        public bool TooglePaneVisibleButton_GetPressed2(IRibbonControl control)
        {
            return provider.TaskPaneIsVisible("2");
        }

        #endregion Public Methods

        #region Private Methods

        private void OnConfigureModuleCatalog(object sender, ProviderEventArgs<IModuleCatalog> e)
        {
            e.Content.AddModule<ExampleView.ExampleModule>();
        }

        private void OnRegisterTypes(object sender, ProviderEventArgs<IContainerRegistry> e)
        {
            e.Content.RegisterInstance<IApplication>(this);
            e.Content.Register<IExampleClass, ExampleClass>();
        }

        #endregion Private Methods
    }
}