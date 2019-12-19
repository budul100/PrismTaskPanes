using System;
using System.Xml.Serialization;

namespace PrismTaskPanes.Settings
{
    [Serializable,
        XmlRoot("PrismTaskPaneSettings")]
    public class PrismTaskPaneSettingsCollection
    {
        #region Public Properties

        [XmlArray("Settings"),
            XmlArrayItem("PrismTaskPane")]
        public PrismTaskPaneSettings[] Settings { get; set; }

        #endregion Public Properties
    }
}