using System;

namespace TestCommon
{
    public class ExampleClass
        : IExampleClass
    {
        #region Public Properties

        public string Message => $"View at {DateTime.Now:fff}.";

        #endregion Public Properties
    }
}