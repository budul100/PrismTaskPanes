using NetOffice.OfficeApi;

namespace PrismTaskPanes.EventArgs
{
    public class TaskPaneEventArgs
        : System.EventArgs
    {
        #region Public Constructors

        public TaskPaneEventArgs(_CustomTaskPane taskPane)
        {
            TaskPane = taskPane;
        }

        #endregion Public Constructors

        #region Public Properties

        public _CustomTaskPane TaskPane { get; }

        #endregion Public Properties
    }
}