using Prism.Regions;
using System;
using System.Runtime.Serialization;

namespace PrismTaskPanes.Exceptions
{
    [Serializable]
    public class RegionNotLoadedException
        : Exception
    {
        #region Public Constructors

        public RegionNotLoadedException(IRegion region)
            : this(GetMessage(region))
        { }

        public RegionNotLoadedException(string message)
            : base(message)
        { }

        public RegionNotLoadedException(string message, Exception innerException)
            : base(message, innerException)
        { }

        public RegionNotLoadedException()
        { }

        #endregion Public Constructors

        #region Protected Constructors

        protected RegionNotLoadedException(SerializationInfo serializationInfo, StreamingContext streamingContext)
        { }

        #endregion Protected Constructors

        #region Private Methods

        private static string GetMessage(IRegion region)
        {
            var result = $"The region {region.Name} could not be loaded.";

            return result;
        }

        #endregion Private Methods
    }
}