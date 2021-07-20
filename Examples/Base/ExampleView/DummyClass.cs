using System;

namespace ExampleView
{
    public class DummyClass
    {
        #region Public Methods

        public static void Dummy<T>()
        {
            static void noop(Type _) { }
            var dummy = typeof(T);
            noop(dummy);
        }

        #endregion Public Methods
    }
}