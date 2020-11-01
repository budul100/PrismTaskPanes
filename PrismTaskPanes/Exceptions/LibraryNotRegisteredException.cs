using System;
using System.Runtime.Serialization;

namespace PrismTaskPanes.Exceptions
{
    [Serializable]
    public class LibraryNotRegisteredException
        : Exception
    {
        #region Public Constructors

        public LibraryNotRegisteredException()
            : base($"The {nameof(PrismTaskPanes)} dll file has not been com registered on the current system.")
        {
        }

        public LibraryNotRegisteredException(string message)
            : base(message)
        {
        }

        public LibraryNotRegisteredException(string message, Exception innerException)
            : base(message, innerException)
        {
        }

        #endregion Public Constructors

        #region Protected Constructors

        protected LibraryNotRegisteredException(SerializationInfo info, StreamingContext context) : base(info, context)
        {
        }

        #endregion Protected Constructors
    }
}