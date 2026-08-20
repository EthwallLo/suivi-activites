using System.Windows;
using System.Windows.Controls;
using MonTableurApp.Models;
using MonTableurApp.ViewModels;

namespace MonTableurApp.Views
{
    public partial class VueArchivesView : UserControl
    {
        public VueArchivesView()
        {
            InitializeComponent();
        }

        private void DesarchiverProjet_Click(object sender, RoutedEventArgs e)
        {
            if ((sender as FrameworkElement)?.DataContext is not Projet projet ||
                DataContext is not MainViewModel viewModel)
            {
                return;
            }

            string nomProjet = string.IsNullOrWhiteSpace(projet.NomProduit)
                ? projet.NumeroProjet ?? "ce projet"
                : projet.NomProduit;

            MessageBoxResult result = MessageBox.Show(
                $"Désarchiver \"{nomProjet}\" et le remettre dans les vues de travail ?",
                "Archivage",
                MessageBoxButton.YesNo,
                MessageBoxImage.Question);

            if (result != MessageBoxResult.Yes)
            {
                return;
            }

            viewModel.ToggleProjetArchive(projet);
        }
    }
}
