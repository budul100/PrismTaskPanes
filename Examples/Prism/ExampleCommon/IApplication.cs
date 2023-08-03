using System;

namespace ExampleCommon
{
    public interface IApplication
    {
        #region Public Methods

        void CallDispatcher(Action callback);

        #endregion Public Methods
    }
}