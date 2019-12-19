using CommonServiceLocator;

namespace PrismTaskPanes.Interfaces
{
    internal interface IOfficeApplication
    {
        #region Public Properties

        IServiceLocator ServiceLocator { get; }

        #endregion Public Properties

        #region Public Methods

        bool GetTaskPaneVisibility(int hash);

        void SetTaskPaneVisibility(int hash, bool isVisible);

        #endregion Public Methods
    }
}