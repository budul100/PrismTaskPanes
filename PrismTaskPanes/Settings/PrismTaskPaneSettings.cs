﻿using NetOffice.OfficeApi.Enums;
using System;
using System.Xml.Serialization;

namespace PrismTaskPanes.Settings
{
    [Serializable]
    [XmlType("PrismTaskPane")]
    public class PrismTaskPaneSettings
    {
        #region Public Properties

        public int AttributeHash { get; set; }

        public MsoCTPDockPosition DockPosition { get; set; } =
            MsoCTPDockPosition.msoCTPDockPositionRight;

        [XmlIgnore]
        public MsoCTPDockPositionRestrict DockRestriction { get; set; } =
            MsoCTPDockPositionRestrict.msoCTPDockPositionRestrictNone;

        public int DocumentHash { get; set; }

        public int Height { get; set; }

        [XmlIgnore]
        public bool InvisibleAtStart { get; set; }

        [XmlIgnore]
        public string NavigationKey { get; set; }

        [XmlIgnore]
        public string NavigationValue { get; set; }

        [XmlIgnore]
        public string RegionContext { get; set; }

        [XmlIgnore]
        public string RegionName { get; set; }

        [XmlIgnore]
        public string Title { get; set; }

        [XmlIgnore]
        public Type View { get; set; }

        public bool Visible { get; set; }

        public int Width { get; set; }

        #endregion Public Properties
    }
}