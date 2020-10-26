using System;

namespace TestCommon
{
    public class TestClass
        : ITestClass
    {
        #region Public Properties

        public string Message => $"View at {DateTime.Now:fff}.";

        #endregion Public Properties
    }
}