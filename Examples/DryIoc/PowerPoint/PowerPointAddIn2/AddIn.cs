#pragma warning disable IDE0060 // Nicht verwendete Parameter entfernen
#pragma warning disable RCS1163 // Unused parameter.
#pragma warning disable RCS1132 // Remove redundant overriding member.

using ExampleCommon;
using ExampleView.Views;
using NetOffice.OfficeApi;
using NetOffice.PowerPointApi.Tools;
using NetOffice.Tools;
using Prism.Ioc;
using Prism.Modularity;
using PrismTaskPanes.Attributes;
using PrismTaskPanes.DryIoc;
using PrismTaskPanes.DryIoc.PowerPoint;
using PrismTaskPanes.EventArgs;
using PrismTaskPanes.Extensions;
using PrismTaskPanes.Interfaces;
using System;
using System.Runtime.InteropServices;
using TestCommon;

namespace PowerPointAddIn2
{
    [COMAddin("PrismTaskPanes.PowerPointAddIn2", "This is an example addin", LoadBehavior.LoadAtStartup),
        ProgId("PowerPointAddIn2.AddIn"),
        Guid("1E6A69D6-D2BC-4D96-8824-4BEC4C1CAE88"),
        RegistryLocation(RegistrySaveLocation.LocalMachine),
        Codebase,
        ComVisible(true)]
    [CustomUI("RibbonUI.xml", true)]
    [PrismTaskPane("1", "ExampleAddin 2 A", typeof(ViewAView), "ExampleRegion", invisibleAtStart: true)]
    [PrismTaskPane("2", "ExampleAddin 2 B", typeof(ViewAView), "ExampleRegion")]
    public class AddIn
        : COMAddin, ITaskPanesReceiver, IApplication
    {
        #region Private Fields

        private TaskPanesProvider provider;

        #endregion Private Fields

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
                ctpFactoryInst: CTPFactoryInst);

            provider.OnConfigureModuleCatalogEvent += OnConfigureModuleCatalog;
            provider.OnRegisterTypesEvent += OnRegisterTypes;
        }

        public override void CustomUI_OnLoad(NetOffice.OfficeApi.Native.IRibbonUI ribbonUI)
        {
            base.CustomUI_OnLoad(ribbonUI);
        }

        public void InvalidateRibbonUI()
        {
            RibbonUI?.Invalidate();
        }

        public void TooglePaneVisibleButton_Click(IRibbonControl control, bool pressed)
        {
            if (provider.CanBeLoaded())
            {
                provider.SetTaskPaneVisibility(
                    id: "1",
                    isVisible: pressed);
            }
        }

        public void TooglePaneVisibleButton_Click2(IRibbonControl control, bool pressed)
        {
            if (provider.CanBeLoaded())
            {
                provider.SetTaskPaneVisibility(
                    id: "2",
                    isVisible: pressed);
            }
        }

        public bool TooglePaneVisibleButton_GetPressed(IRibbonControl control)
        {
            if (!provider.CanBeLoaded())
            {
                return false;
            }

            return provider.TaskPaneIsVisible(
                id: "1");
        }

        public bool TooglePaneVisibleButton_GetPressed2(IRibbonControl control)
        {
            if (!provider.CanBeLoaded())
            {
                return false;
            }

            return provider.TaskPaneIsVisible(
                id: "1");
        }

        #endregion Public Methods

        #region Private Methods

        private void Addin_OnStartupComplete(ref Array custom)
        {
            Console.WriteLine($"Addin started in Excel Version {Application.Version}");
        }

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