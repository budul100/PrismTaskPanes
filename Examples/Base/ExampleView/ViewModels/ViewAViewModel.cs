using DryIoc;
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

        private string _message;

        #endregion Private Fields

        #region Public Constructors

        public ViewAViewModel(IResolverContext container, IExampleClass test)
        {
            TestCommand = new DelegateCommand(TestAction);

            var test2 = container.Resolve<IExampleClass>();

            Message = $"{test.Message} - {test2.Message}";
        }

        #endregion Public Constructors

        #region Public Properties

        public string Message
        {
            get { return _message; }
            set { SetProperty(ref _message, value); }
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