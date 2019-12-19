using Prism.Regions;
using System;
using System.Collections;
using System.Collections.Specialized;
using System.Windows;

namespace PrismTaskPanes.Regions
{
    internal class DisposableRegionBehavior :
         RegionBehavior
    {
        #region Public Properties

        public static string BehaviorKey => "DisposableRegion";

        #endregion Public Properties

        #region Protected Methods

        protected override void OnAttach()
        {
            Region.Views.CollectionChanged += OnActiveViewsChanged;
        }

        #endregion Protected Methods

        #region Private Methods

        private static void DisposeViewsOrViewModels(IList views)
        {
            foreach (var view in views)
            {
                if (view is FrameworkElement frameworkElement)
                {
                    if (frameworkElement.DataContext is IDisposable disposableDataContext)
                        disposableDataContext.Dispose();

                    if (view is IDisposable)
                    {
                        var disposableView = view as IDisposable;
                        disposableView.Dispose();
                    }
                }
            }
        }

        private static void OnActiveViewsChanged(object sender, NotifyCollectionChangedEventArgs e)
        {
            if (e.Action == NotifyCollectionChangedAction.Remove)
            {
                DisposeViewsOrViewModels(e.OldItems);
            }
        }

        #endregion Private Methods
    }
}