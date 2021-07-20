using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace PrismTaskPanes.Controls
{
    [Guid("EFB770FC-6A2A-4D19-8AF8-5856F2BE6C34"),
        ComDefaultInterface(typeof(IPrismTaskPanesHost)),
        ComVisible(true)]
    public partial class PrismTaskPanesHost
        : UserControl, IPrismTaskPanesHost
    {
        #region Public Constructors

        public PrismTaskPanesHost()
        {
            InitializeComponent();
        }

        #endregion Public Constructors
    }
}