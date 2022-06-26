using System;
using System.Runtime.InteropServices;
using System.Runtime.Serialization;

namespace PrismTaskPanes.Exceptions
{
    [ComVisible(false)]
    [Serializable]
    public class HostNotRegisteredException
        : Exception
    {
        #region Public Constructors

        public HostNotRegisteredException()
            : this(GetMessage())
        { }

        public HostNotRegisteredException(string message)
            : base(message)
        { }

        public HostNotRegisteredException(string message, Exception innerException)
            : base(message, innerException)
        { }

        #endregion Public Constructors

        #region Protected Constructors

        protected HostNotRegisteredException(SerializationInfo info, StreamingContext context) : base(info, context)
        { }

        #endregion Protected Constructors

        #region Private Methods

        private static string GetMessage()
        {
            var result = $"The {nameof(PrismTaskPanes)} dll file has not been com registered on the current system.";

            return result;
        }

        #endregion Private Methods
    }
}