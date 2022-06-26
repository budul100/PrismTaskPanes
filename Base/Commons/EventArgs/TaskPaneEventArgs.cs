using System.Runtime.InteropServices;

namespace PrismTaskPanes.EventArgs
{
    [ComVisible(false)]
    public class TaskPaneEventArgs
        : System.EventArgs
    {
        #region Public Constructors

        public TaskPaneEventArgs(object taskPane)
        {
            TaskPane = taskPane;
        }

        #endregion Public Constructors

        #region Public Properties

        public object TaskPane { get; }

        #endregion Public Properties
    }
}