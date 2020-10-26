using NetOffice.OfficeApi.Enums;
using PrismTaskPanes.Interfaces;
using System;
using System.ComponentModel.Composition;

namespace PrismTaskPanes.Attributes
{
    [AttributeUsage(AttributeTargets.Class, AllowMultiple = true)]
    public sealed class PrismTaskPaneAttribute
        : ExportAttribute
    {
        #region Public Constructors

        public PrismTaskPaneAttribute()
            : base(typeof(ITaskPanesReceiver))
        { }

        public PrismTaskPaneAttribute(
            string id,
            string title,
            Type view,
            string regionName,
            bool visible = false,
            bool invisibleAtStart = false,
            int width = 0,
            int height = 0,
            MsoCTPDockPosition dockPosition = MsoCTPDockPosition.msoCTPDockPositionRight,
            MsoCTPDockPositionRestrict dockRestriction = MsoCTPDockPositionRestrict.msoCTPDockPositionRestrictNone) :
            this()
        {
            if (string.IsNullOrWhiteSpace(id))
            {
                throw new ArgumentNullException(nameof(id));
            }

            if (string.IsNullOrWhiteSpace(title))
            {
                throw new ArgumentNullException(nameof(title));
            }

            if (string.IsNullOrWhiteSpace(regionName))
            {
                throw new ArgumentNullException(nameof(regionName));
            }

            View = view ?? throw new ArgumentNullException(nameof(view));

            ID = id;
            Title = title;
            RegionName = regionName;

            Visible = visible;
            InvisibleAtStart = invisibleAtStart;
            DockPosition = dockPosition;
            DockRestriction = dockRestriction;

            if (width > 0
                && DockPosition != MsoCTPDockPosition.msoCTPDockPositionTop
                && DockPosition != MsoCTPDockPosition.msoCTPDockPositionBottom)
                Width = width;
            if (height > 0
                && DockPosition != MsoCTPDockPosition.msoCTPDockPositionLeft
                && DockPosition != MsoCTPDockPosition.msoCTPDockPositionRight)
                Height = height;
        }

        public PrismTaskPaneAttribute(
            string id,
            string title,
            Type view,
            string regionName,
            string navigationKey,
            string navigationValue,
            bool visible = false,
            bool invisibleAtStart = false,
            int width = 0,
            int height = 0,
            MsoCTPDockPosition dockPosition = MsoCTPDockPosition.msoCTPDockPositionRight,
            MsoCTPDockPositionRestrict dockRestriction = MsoCTPDockPositionRestrict.msoCTPDockPositionRestrictNone) :
            this(id: id, title: title, view: view, regionName: regionName, visible: visible,
                invisibleAtStart: invisibleAtStart, width: width, height: height, dockPosition: dockPosition,
                dockRestriction: dockRestriction)
        {
            if (string.IsNullOrWhiteSpace(navigationKey))
            {
                throw new ArgumentNullException(nameof(navigationKey));
            }

            NavigationKey = navigationKey;
            NavigationValue = navigationValue;
        }

        public PrismTaskPaneAttribute(
            string id,
            string title,
            Type view,
            string regionName,
            string regionContext,
            bool visible = false,
            bool invisibleAtStart = false,
            int width = 0,
            int height = 0,
            MsoCTPDockPosition dockPosition = MsoCTPDockPosition.msoCTPDockPositionRight,
            MsoCTPDockPositionRestrict dockRestriction = MsoCTPDockPositionRestrict.msoCTPDockPositionRestrictNone) :
            this(id: id, title: title, view: view, regionName: regionName, visible: visible,
                invisibleAtStart: invisibleAtStart, width: width, height: height, dockPosition: dockPosition,
                dockRestriction: dockRestriction)
        {
            if (string.IsNullOrWhiteSpace(regionContext))
            {
                throw new ArgumentNullException(nameof(regionContext));
            }

            RegionContext = regionContext;
        }

        #endregion Public Constructors

        #region Public Properties

        public MsoCTPDockPosition DockPosition { get; set; } =
            MsoCTPDockPosition.msoCTPDockPositionRight;

        public MsoCTPDockPositionRestrict DockRestriction { get; set; } =
            MsoCTPDockPositionRestrict.msoCTPDockPositionRestrictNone;

        public int Height { get; set; }

        public string ID { get; set; }

        public bool InvisibleAtStart { get; set; }

        public string NavigationKey { get; set; }

        public string NavigationValue { get; set; }

        public string ReceiverHash { get; set; }

        public string RegionContext { get; set; }

        public string RegionName { get; set; }

        public string Title { get; set; }

        public Type View { get; set; }

        public bool Visible { get; set; }

        public int Width { get; set; }

        #endregion Public Properties
    }
}