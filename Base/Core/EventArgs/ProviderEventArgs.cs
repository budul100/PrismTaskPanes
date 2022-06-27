using System.Runtime.InteropServices;

namespace PrismTaskPanes.EventArgs
{
    [ComVisible(false)]
    public class ProviderEventArgs<T>
        : System.EventArgs
    {
        #region Public Constructors

        public ProviderEventArgs(T content)
        {
            Content = content;
        }

        #endregion Public Constructors

        #region Public Properties

        public T Content { get; }

        #endregion Public Properties
    }
}