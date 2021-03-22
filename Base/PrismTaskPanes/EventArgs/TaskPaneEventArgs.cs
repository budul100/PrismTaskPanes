using NetOffice.OfficeApi;
using System.Runtime.InteropServices;

namespace PrismTaskPanes.EventArgs
{
    [ComVisible(false)]
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