using System;
using System.Runtime.InteropServices;
using System.Runtime.Serialization;

namespace PrismTaskPanes.Exceptions
{
    [ComVisible(false)]
    [Serializable]
    public class AlreadyLoadedException
        : Exception
    {
        #region Public Fields

        public const string Text = "There has been a PrismTaskPane based Add-In loaded already. " +
            "Currently there can't be multiple Add-Ins opened concurrently in the same Office " +
            "application session. To open this Add-In, please restart the respective Office application " +
            "and try to open the Add-In again.";

        #endregion Public Fields

        #region Public Constructors

        public AlreadyLoadedException()
            : this(Text)
        { }

        public AlreadyLoadedException(string message)
            : base(message)
        { }

        public AlreadyLoadedException(string message, Exception innerException)
            : base(message, innerException)
        { }

        #endregion Public Constructors

        #region Protected Constructors

        protected AlreadyLoadedException(SerializationInfo info, StreamingContext context)
            : base(info, context)
        { }

        #endregion Protected Constructors
    }
}