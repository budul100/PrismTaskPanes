using DryIoc;
using Prism.Regions;
using PrismTaskPanes;
using TestCommon;

namespace TestViewLib.ViewModels
{
    public class ViewAViewModel
        : ScopedContentBase
    {
        #region Private Fields

        private string _message;

        #endregion Private Fields

        #region Public Constructors

        public ViewAViewModel()
        {
            var test = DryIocProvider.ResolverContext.Resolve<ITestInterface>();
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