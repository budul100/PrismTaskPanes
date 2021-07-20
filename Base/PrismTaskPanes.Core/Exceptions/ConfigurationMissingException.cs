using System;
using System.Runtime.Serialization;

namespace PrismTaskPanes.Exceptions
{
    public class ConfigurationMissingException
        : Exception
    {
        #region Public Constructors

        public ConfigurationMissingException()
            : this("There is no respective Prism Task Pane defined.")
        {
        }

        public ConfigurationMissingException(string message)
            : base(message)
        {
        }

        public ConfigurationMissingException(string message, Exception innerException)
            : base(message, innerException)
        {
        }

        #endregion Public Constructors

        #region Protected Constructors

        protected ConfigurationMissingException(SerializationInfo info, StreamingContext context)
            : base(info, context)
        {
        }

        #endregion Protected Constructors
    }
}