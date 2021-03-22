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
            : this(GetMessage())
        { }

        public LibraryNotRegisteredException(string message)
            : base(message)
        { }

        public LibraryNotRegisteredException(string message, Exception innerException)
            : base(message, innerException)
        { }

        #endregion Public Constructors

        #region Protected Constructors

        protected LibraryNotRegisteredException(SerializationInfo info, StreamingContext context) : base(info, context)
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