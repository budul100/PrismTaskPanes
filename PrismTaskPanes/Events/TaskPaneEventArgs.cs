using NetOffice.OfficeApi;
using System;

namespace PrismTaskPanes.Events
{
    public class TaskPaneEventArgs
        : EventArgs
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