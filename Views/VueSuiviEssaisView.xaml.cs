using System.Windows;
using System.Windows.Controls;
using MonTableurApp.ViewModels;

namespace MonTableurApp.Views
{
    public partial class VueSuiviEssaisView : UserControl
    {
        public VueSuiviEssaisView()
        {
            InitializeComponent();
        }

        private void EssaisATraiter_Click(object sender, RoutedEventArgs e)
        {
            if (DataContext is MainViewModel viewModel)
            {
                viewModel.SetEssaiFilterToToProcess();
            }
        }

        private void EssaisEnCours_Click(object sender, RoutedEventArgs e)
        {
            if (DataContext is MainViewModel viewModel)
            {
                viewModel.SetEssaiFilterToInProgress();
            }
        }

        private void EssaisTermines_Click(object sender, RoutedEventArgs e)
        {
            if (DataContext is MainViewModel viewModel)
            {
                viewModel.SetEssaiFilterToDone();
            }
        }
    }
}
