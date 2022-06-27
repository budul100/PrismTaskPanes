using System.Runtime.InteropServices;

namespace PrismTaskPanes.Enums
{
    [ComVisible(false)]
    public enum DockRestriction
    {
        None,

        NoHorizontal,

        NoVertical,

        NoChange,
    }
}