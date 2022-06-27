using PrismTaskPanes.Extensions;
using System.Windows;
using System.Windows.Controls;

namespace ExampleView.Views
{
    /// <summary>
    /// Interaction logic for ViewA.xaml
    /// </summary>
    public partial class ViewAView : UserControl
    {
        #region Public Constructors

        public ViewAView()
        {
            // This pre-load is necessary for the trigger in ViewA only
            // See https://github.com/microsoft/XamlBehaviorsWpf/issues/86

            var _ = new Microsoft.Xaml.Behaviors.DefaultTriggerAttribute(
                targetType: typeof(Trigger),
                triggerType: typeof(Microsoft.Xaml.Behaviors.TriggerBase),
                parameters: null);

            this.LoadViewFromUri("/ExampleView;component/views/viewaview.xaml");

            InitializeComponent();
        }

        #endregion Public Constructors
    }
}