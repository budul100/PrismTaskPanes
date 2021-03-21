using NetOffice.PowerPointApi;
using Prism.Ioc;
using System;
using System.IO;

namespace PrismTaskPanes.Applications.DryIoc
{
    internal class PowerPointApplication
        : OfficeApplication, IDisposable
    {
        #region Private Fields

        private readonly Application application;

        #endregion Private Fields

        #region Public Constructors

        public PowerPointApplication(object application, object ctpFactoryInst)
            : base(application, ctpFactoryInst)
        {
            this.application = application as Application;

            this.application.NewPresentationEvent += OnElementNew; ;
            this.application.PresentationOpenEvent += OnElementOpen;
            this.application.PresentationSaveEvent += OnElementSaveAfter;
            this.application.PresentationBeforeCloseEvent += OnElementCloseBefore;

            this.application.OnDispose += OnApplicationDispose;
        }

        #endregion Public Constructors

        #region Protected Properties

        protected override object TaskPaneWindow => GetActiveWindow();

        protected override int? TaskPaneWindowKey => GetActiveWindow()?.HWND;

        #endregion Protected Properties

        #region Protected Methods

        protected override void Dispose(bool disposing)
        {
            if (!isDisposed)
            {
                if (disposing)
                {
                    application?.DisposeChildInstances();
                }

                base.Dispose(disposing);
            }
        }

        protected override string GetTaskPaneIdentifier()
        {
            var path = GetActiveElement()?.FullName;
            var isPath = !string.IsNullOrWhiteSpace(Path.GetDirectoryName(path));

            return isPath
                ? path
                : default;
        }

        protected override void RegisterTypes(IContainerRegistry containerRegistry)
        {
            base.RegisterTypes(containerRegistry);

            containerRegistry.RegisterInstance(application);
        }

        #endregion Protected Methods

        #region Private Methods

        private Presentation GetActiveElement()
        {
            var result = default(Presentation);

            try
            {
                result = application?.ActivePresentation;
            }
            catch (NetOffice.Exceptions.PropertyGetCOMException)
            { }

            return result;
        }

        private DocumentWindow GetActiveWindow()
        {
            var result = default(DocumentWindow);

            try
            {
                result = application?.ActiveWindow;
            }
            catch (NetOffice.Exceptions.PropertyGetCOMException)
            { }

            return result;
        }

        private void OnElementCloseBefore(Presentation pres, ref bool cancel)
        {
            SaveScope();
        }

        private void OnElementNew(Presentation pres)
        {
            OpenScope();
        }

        private void OnElementOpen(Presentation pres)
        {
            OpenScope();
        }

        private void OnElementSaveAfter(Presentation pres)
        {
            SaveScope();
        }

        #endregion Private Methods
    }
}