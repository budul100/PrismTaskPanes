using DryIoc;
using System;

namespace TestCommon
{
    public class ExampleClass
        : IExampleClass
    {
        #region Private Fields

        private readonly IContainer container;

        #endregion Private Fields

        #region Public Constructors

        public ExampleClass()
        { }

        public ExampleClass(IContainer container)
        {
            this.container = container;
        }

        #endregion Public Constructors

        #region Public Properties

        public string ContainerMessage => container.GetHashCode().ToString();

        public string Message => $"View at {DateTime.Now:fff}.";

        #endregion Public Properties
    }
}