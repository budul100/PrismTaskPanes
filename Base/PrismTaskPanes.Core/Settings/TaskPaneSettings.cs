using PrismTaskPanes.Commons.Enums;
using System;
using System.Runtime.InteropServices;
using System.Xml.Serialization;

namespace PrismTaskPanes.Settings
{
    [ComVisible(false)]
    [Serializable]
    [XmlType("PrismTaskPaneConfiguration")]
    public class TaskPaneSettings
    {
        #region Public Properties

        public DockPosition DockPosition { get; set; } = DockPosition.Right;

        [XmlIgnore]
        public DockRestriction DockRestriction { get; set; } = DockRestriction.None;

        public string DocumentHash { get; set; }

        public int Height { get; set; }

        [XmlIgnore]
        public bool InvisibleAtStart { get; set; }

        [XmlIgnore]
        public string NavigationValue { get; set; }

        public string ReceiverHash { get; set; }

        [XmlIgnore]
        public string RegionContext { get; set; }

        [XmlIgnore]
        public string RegionName { get; set; }

        [XmlIgnore]
        public ScrollVisibility ScrollBarHorizontal { get; set; } = ScrollVisibility.Auto;

        [XmlIgnore]
        public ScrollVisibility ScrollBarVertical { get; set; } = ScrollVisibility.Auto;

        [XmlIgnore]
        public string Title { get; set; }

        [XmlIgnore]
        public Type View { get; set; }

        [XmlIgnore]
        public Uri ViewUri { get; set; }

        public bool Visible { get; set; }

        public int Width { get; set; }

        #endregion Public Properties
    }
}