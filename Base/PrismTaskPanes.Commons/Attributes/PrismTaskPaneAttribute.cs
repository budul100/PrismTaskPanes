using PrismTaskPanes.Commons.Enums;
using System;
using System.Runtime.InteropServices;

namespace PrismTaskPanes.Attributes
{
    [ComVisible(false)]
    [AttributeUsage(AttributeTargets.Class, AllowMultiple = true)]
    public sealed class PrismTaskPaneAttribute
        : Attribute
    {
        #region Public Constructors

        public PrismTaskPaneAttribute(
            string id,
            string title,
            Type view,
            string regionName,
            bool visible = false,
            bool invisibleAtStart = false,
            int width = 0,
            int height = 0,
            DockPosition dockPosition = DockPosition.Right,
            DockRestriction dockRestriction = DockRestriction.None,
            ScrollVisibility scrollBarHorizontal = ScrollVisibility.Auto,
            ScrollVisibility scrollBarVertical = ScrollVisibility.Auto)
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
            ScrollBarHorizontal = scrollBarHorizontal;
            ScrollBarVertical = scrollBarVertical;

            if (width > 0
                && DockPosition != DockPosition.Top
                && DockPosition != DockPosition.Bottom)
                Width = width;
            if (height > 0
                && DockPosition != DockPosition.Left
                && DockPosition != DockPosition.Right)
                Height = height;
        }

        public PrismTaskPaneAttribute(
            string id,
            string title,
            Type view,
            string regionName,
            string navigationValue,
            bool visible = false,
            bool invisibleAtStart = false,
            int width = 0,
            int height = 0,
            DockPosition dockPosition = DockPosition.Right,
            DockRestriction dockRestriction = DockRestriction.None,
            ScrollVisibility scrollBarHorizontal = ScrollVisibility.Auto,
            ScrollVisibility scrollBarVertical = ScrollVisibility.Auto) :
            this(id: id, title: title, view: view, regionName: regionName, visible: visible, invisibleAtStart: invisibleAtStart,
                width: width, height: height, dockPosition: dockPosition, dockRestriction: dockRestriction,
                scrollBarHorizontal: scrollBarHorizontal, scrollBarVertical: scrollBarVertical)
        {
            if (string.IsNullOrWhiteSpace(navigationValue))
            {
                throw new ArgumentNullException(nameof(navigationValue));
            }

            NavigationValue = navigationValue;
        }

        public PrismTaskPaneAttribute(
            string id,
            string title,
            Type view,
            string regionName,
            string navigationValue,
            string regionContext,
            bool visible = false,
            bool invisibleAtStart = false,
            int width = 0,
            int height = 0,
            DockPosition dockPosition = DockPosition.Right,
            DockRestriction dockRestriction = DockRestriction.None,
            ScrollVisibility scrollBarHorizontal = ScrollVisibility.Auto,
            ScrollVisibility scrollBarVertical = ScrollVisibility.Auto) :
            this(id: id, title: title, view: view, regionName: regionName, navigationValue: navigationValue, visible: visible,
                invisibleAtStart: invisibleAtStart, width: width, height: height, dockPosition: dockPosition,
                dockRestriction: dockRestriction, scrollBarHorizontal: scrollBarHorizontal, scrollBarVertical: scrollBarVertical)
        {
            if (string.IsNullOrWhiteSpace(regionContext))
            {
                throw new ArgumentNullException(nameof(regionContext));
            }

            RegionContext = regionContext;
        }

        #endregion Public Constructors

        #region Public Properties

        public DockPosition DockPosition { get; set; } = DockPosition.Right;

        public DockRestriction DockRestriction { get; set; } = DockRestriction.None;

        public int Height { get; set; }

        public string ID { get; set; }

        public bool InvisibleAtStart { get; set; }

        public string NavigationValue { get; set; }

        public string ReceiverHash { get; set; }

        public string RegionContext { get; set; }

        public string RegionName { get; set; }

        public ScrollVisibility ScrollBarHorizontal { get; set; } = ScrollVisibility.Auto;

        public ScrollVisibility ScrollBarVertical { get; set; } = ScrollVisibility.Auto;

        public string Title { get; set; }

        public Type View { get; set; }

        public bool Visible { get; set; }

        public int Width { get; set; }

        #endregion Public Properties
    }
}