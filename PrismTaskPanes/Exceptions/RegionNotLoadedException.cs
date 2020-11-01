﻿using Prism.Regions;
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
            : base(GetMessage(region))
        {
        }

        public RegionNotLoadedException(IRegion region, Exception innerException)
            : base(GetMessage(region), innerException)
        {
        }

        public RegionNotLoadedException(string message)
            : base(message)
        {
        }

        public RegionNotLoadedException(string message, Exception innerException)
            : base(message, innerException)
        {
        }

        public RegionNotLoadedException()
        {
        }

        #endregion Public Constructors

        #region Protected Constructors

        protected RegionNotLoadedException(SerializationInfo serializationInfo, StreamingContext streamingContext)
        {
            throw new NotImplementedException();
        }

        #endregion Protected Constructors

        #region Private Methods

        private static string GetMessage(IRegion region)
        {
            var result = $"The region {region.Name} could not be loaded;";

            return result;
        }

        #endregion Private Methods
    }
}