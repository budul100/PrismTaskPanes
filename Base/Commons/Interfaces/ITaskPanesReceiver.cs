using System.Runtime.InteropServices;

namespace PrismTaskPanes.Interfaces
{
    [ComVisible(false)]
    public interface ITaskPanesReceiver
    {
        #region Public Methods

        void InvalidateRibbonUI();

        #endregion Public Methods
    }
}