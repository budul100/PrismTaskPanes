using PrismTaskPanes.Enums;
using System;
using System.Runtime.InteropServices;

namespace PrismTaskPanes.Interfaces
{
    [ComVisible(false)]
    public interface IPrismTaskPanesView
    {
        #region Public Methods

        void Initialize(string regionName, object regionContext, Uri viewUri,
            ScrollVisibility scrollBarHorizontal, ScrollVisibility scrollBarVertical);

        #endregion Public Methods
    }
}