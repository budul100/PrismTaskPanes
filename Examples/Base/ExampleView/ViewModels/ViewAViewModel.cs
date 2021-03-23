using Prism.Commands;
using Prism.Regions;
using System;
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

        public ViewAViewModel(IExampleClass test)
        {
            TestCommand = new DelegateCommand(TestAction);

            Message = test.Message;
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
            Message = DateTime.Now.ToLongTimeString();
        }

        #endregion Private Methods
    }
}