using DryIoc;

namespace PrismTaskPanes.DryIoc.EventArgs
{
    public class ExcelEventArgs
        : System.EventArgs
    {
        #region Public Constructors

        public ExcelEventArgs(IResolverContext container)
        {
            Container = container;
        }

        #endregion Public Constructors

        #region Public Properties

        public IResolverContext Container { get; }

        #endregion Public Properties
    }
}