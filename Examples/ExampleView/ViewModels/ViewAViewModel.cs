﻿using Prism.Regions;
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
            Message = test.Message;
        }

        #endregion Public Constructors

        #region Public Properties

        public string Message
        {
            get { return _message; }
            set { SetProperty(ref _message, value); }
        }

        #endregion Public Properties

        #region Protected Methods

        protected override void Initialize(NavigationContext navigationContext)
        {
        }

        #endregion Protected Methods
    }
}