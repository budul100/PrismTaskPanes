using CommonServiceLocator;

namespace PrismTaskPanes.Interfaces
{
    internal interface IOfficeApplication
    {
        #region Public Properties

        IServiceLocator ServiceLocator { get; }

        #endregion Public Properties

        #region Public Methods

        bool IsTaskPaneExist(int hash);

        bool IsTaskPaneVisible(int hash);

        void SetTaskPaneVisibility(int hash, bool isVisible);

        #endregion Public Methods
    }
}