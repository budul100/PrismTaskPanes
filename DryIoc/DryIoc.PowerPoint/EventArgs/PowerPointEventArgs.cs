using DryIoc;

namespace PrismTaskPanes.DryIoc.EventArgs
{
    public class PowerPointEventArgs
        : System.EventArgs
    {
        #region Public Constructors

        public PowerPointEventArgs(IResolverContext container)
        {
            Container = container;
        }

        #endregion Public Constructors

        #region Public Properties

        public IResolverContext Container { get; }

        #endregion Public Properties
    }
}