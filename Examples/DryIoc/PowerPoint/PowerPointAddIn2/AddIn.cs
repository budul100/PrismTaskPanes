﻿using ExampleView.Views;
using NetOffice.OfficeApi;
using NetOffice.PowerPointApi.Tools;
using NetOffice.Tools;
using Prism.Ioc;
using Prism.Modularity;
using PrismTaskPanes.Attributes;
using PrismTaskPanes.DryIoc;
using PrismTaskPanes.Interfaces;
using System;
using System.Runtime.InteropServices;
using TestCommon;

namespace PowerPointAddIn2
{
    [COMAddin("PrismTaskPanes.PowerPointAddIn2", "This is an example addin", LoadBehavior.LoadAtStartup),
        ProgId("PowerPointAddIn2.AddIn"),
        Guid("1E6A69D6-D2BC-4D96-8824-4BEC4C1CAE88"),
        RegistryLocation(RegistrySaveLocation.CurrentUser)]
    [CustomUI("RibbonUI.xml", true)]
    [PrismTaskPane("1", "ExampleAddin 2 A", typeof(ViewAView), "ExampleRegion", invisibleAtStart: true)]
    [PrismTaskPane("2", "ExampleAddin 2 B", typeof(ViewAView), "ExampleRegion")]
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
            moduleCatalog.AddModule<ExampleView.Module>(nameof(PowerPointAddIn2));
        }

        public override void CTPFactoryAvailable(object CTPFactoryInst)
        {
            base.CTPFactoryAvailable(CTPFactoryInst);

            this.InitializeProvider(
                application: Application,
                ctpFactoryInst: CTPFactoryInst);
        }

        public override void CustomUI_OnLoad(NetOffice.OfficeApi.Native.IRibbonUI ribbonUI)
        {
            base.CustomUI_OnLoad(ribbonUI);
        }

        public void InvalidateRibbonUI()
        {
            RibbonUI?.Invalidate();
        }

        public void RegisterTypes(IContainerRegistry builder)
        {
            builder.Register<IExampleClass, ExampleClass>();
        }

        public void RegisterTypes(IContainerProvider containerProvider)
        { }

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
            Console.WriteLine($"Addin started in Excel Version {Application.Version}");
        }

        #endregion Private Methods
    }
}