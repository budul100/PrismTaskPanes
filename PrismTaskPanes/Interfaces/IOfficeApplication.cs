using DryIoc;

namespace PrismTaskPanes.Interfaces
{
    internal interface IOfficeApplication
    {
        #region Public Methods

        IScope GetCurrentScope();

        void SetTaskPaneVisible(int hash, bool isVisible);

        bool TaskPaneExists(int hash);

        bool TaskPaneVisible(int hash);

        #endregion Public Methods
    }
}