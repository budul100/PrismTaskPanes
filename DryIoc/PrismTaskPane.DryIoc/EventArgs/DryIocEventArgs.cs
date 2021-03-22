using DryIoc;

namespace PrismTaskPanes.DryIoc.EventArgs
{
    public class DryIocEventArgs
        : System.EventArgs
    {
        #region Public Constructors

        public DryIocEventArgs(IResolverContext container)
        {
            Container = container;
        }

        #endregion Public Constructors

        #region Public Properties

        public IResolverContext Container { get; }

        #endregion Public Properties
    }
}