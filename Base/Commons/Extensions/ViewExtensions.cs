using System;
using System.IO.Packaging;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Markup;
using System.Windows.Navigation;

namespace PrismTaskPanes.Extensions
{
    [ComVisible(false)]
    public static class ViewExtensions
    {
        #region Private Fields

        private static readonly Uri packageBaseUri = GetPackageBaseUri();

        #endregion Private Fields

        #region Public Methods

        public static void LoadViewFromUri(this UserControl userControl, string baseUri)
        {
            var locator = new Uri(
                uriString: baseUri,
                uriKind: UriKind.Relative);

            var parserUri = new Uri(
                baseUri: packageBaseUri,
                relativeUri: locator);

            var parserContext = new ParserContext
            {
                BaseUri = parserUri
            };

            var exprCa = (PackagePart)typeof(Application)
                .GetMethod(
                    name: "GetResourceOrContentPart",
                    bindingAttr: BindingFlags.NonPublic | BindingFlags.Static)
                .Invoke(
                    obj: null,
                    parameters: new object[] { locator });

            var stream = exprCa.GetStream();

            var parameters = new object[]
            {
                stream,
                parserContext,
                userControl,
                true
            };

            typeof(XamlReader)
                .GetMethod(
                    name: "LoadBaml",
                    bindingAttr: BindingFlags.NonPublic | BindingFlags.Static)
                .Invoke(
                    obj: null,
                    parameters: parameters);
        }

        #endregion Public Methods

        #region Private Methods

        private static Uri GetPackageBaseUri()
        {
            var result = (Uri)typeof(BaseUriHelper)
                .GetProperty(
                    name: "PackAppBaseUri",
                    bindingAttr: BindingFlags.Static | BindingFlags.NonPublic)
                .GetValue(
                    obj: null,
                    index: null);

            return result;
        }

        #endregion Private Methods
    }
}