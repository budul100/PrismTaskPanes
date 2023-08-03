using DryIoc;
using ExampleCommon;
using Prism.Commands;
using Prism.Regions;
using System;
using System.Text;
using TestCommon;

namespace ExampleView.ViewModels
{
    public class ViewAViewModel
        : ScopedContentBase
    {
        #region Private Fields

        private readonly IApplication application;

        private string _message;

        #endregion Private Fields

        #region Public Constructors

        public ViewAViewModel(IResolverContext container, IExampleClass test)
        {
            TestCommand = new DelegateCommand(TestAction);

            this.application = container.Resolve<IApplication>();
            var test2 = container.Resolve<IExampleClass>();

            Message = $"{typeof(ViewAViewModel).Name}\r\n" +
                $"Test 1: {test.Message} / {test.ContainerMessage}\r\n" +
                $"Test 2: {test2.Message} / {test2.ContainerMessage}\r\n" +
                $"Container: {container.GetHashCode()}";
        }

        #endregion Public Constructors

        #region Public Properties

        public string Message
        {
            get { return _message; }
            set
            {
                SetProperty(ref _message, value);

                // This additional callback can be used if the view is run in an asynchronous application
                application.CallDispatcher(() => RaisePropertyChanged(nameof(Message)));
            }
        }

        public DelegateCommand TestCommand { get; }

        #endregion Public Properties

        #region Protected Methods

        protected override void Initialize(NavigationContext navigationContext)
        {
        }

        #endregion Protected Methods

        #region Private Methods

        private void TestAction()
        {
            var message = new StringBuilder();

            for (int i = 0; i < 300; i++)
            {
                message.AppendLine(DateTime.Now.ToLongTimeString());
            }

            Message = message.ToString();
        }

        #endregion Private Methods
    }
}