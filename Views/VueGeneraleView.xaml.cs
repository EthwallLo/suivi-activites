using System.Windows;
using System.Windows.Controls;
using MonTableurApp.ViewModels;

namespace MonTableurApp.Views
{
    public partial class VueGeneraleView : UserControl
    {
        public VueGeneraleView()
        {
            InitializeComponent();
        }

        private void TousProjets_Click(object sender, RoutedEventArgs e)
        {
            if (DataContext is MainViewModel viewModel)
            {
                viewModel.SetQuickFilterToAll();
            }
        }

        private void EnCours_Click(object sender, RoutedEventArgs e)
        {
            if (DataContext is MainViewModel viewModel)
            {
                viewModel.SetQuickFilterToInProgress();
            }
        }

        private void Rapports_Click(object sender, RoutedEventArgs e)
        {
            if (DataContext is MainViewModel viewModel)
            {
                viewModel.SetQuickFilterToReports();
            }
        }

        private void Faits_Click(object sender, RoutedEventArgs e)
        {
            if (DataContext is MainViewModel viewModel)
            {
                viewModel.SetQuickFilterToDone();
            }
        }
    }
}
