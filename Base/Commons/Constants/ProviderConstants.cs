using System.Runtime.InteropServices;

namespace PrismTaskPanes.Constants
{
    [ComVisible(false)]
    public static class ProviderConstants
    {
        #region Private Fields

        private const string navigationKey = "Navigation";
        private const string windowKey = "Window";

        #endregion Private Fields

        #region Public Properties

        public static string NavigationKey => navigationKey;

        public static string WindowKey => windowKey;

        #endregion Public Properties
    }
}