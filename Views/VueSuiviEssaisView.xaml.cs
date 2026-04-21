using System.Windows;
using System.Windows.Controls;
using MonTableurApp.Models;
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

        private void ModifierProduit_Click(object sender, RoutedEventArgs e)
        {
            if (DataContext is not MainViewModel viewModel)
            {
                return;
            }

            if ((sender as FrameworkElement)?.DataContext is not Projet projet)
            {
                return;
            }

            var window = new EditProjetWindow(viewModel, projet)
            {
                Owner = Window.GetWindow(this)
            };

            window.ShowDialog();
        }
    }
}
