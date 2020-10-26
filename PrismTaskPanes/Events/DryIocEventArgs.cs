using DryIoc;
using System;

namespace PrismTaskPanes.Events
{
    public class DryIocEventArgs
        : EventArgs
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