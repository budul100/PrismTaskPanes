namespace PrismTaskPanes.Interfaces
{
    internal interface IOfficeApplication
    {
        #region Public Methods

        bool GetTaskPaneVisibility(int hash);

        void SetTaskPaneVisibility(int hash, bool isVisible);

        #endregion Public Methods
    }
}