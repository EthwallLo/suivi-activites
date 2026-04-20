using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using Microsoft.Win32;
using MonTableurApp.Models;
using MonTableurApp.Services;
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

        private void ExporterTableau_Click(object sender, RoutedEventArgs e)
        {
            if (DataContext is not MainViewModel viewModel)
            {
                return;
            }

            var dialog = new SaveFileDialog
            {
                Title = "Exporter le tableau",
                Filter = "Classeur Excel (*.xlsx)|*.xlsx",
                DefaultExt = ".xlsx",
                AddExtension = true,
                FileName = "suivi-projets.xlsx"
            };

            if (dialog.ShowDialog() != true)
            {
                return;
            }

            List<Projet> projets = viewModel.ProjetsView.Cast<Projet>().ToList();
            XlsxExportService.ExportProjects(dialog.FileName, projets);

            MessageBox.Show(
                $"Export terminé.\n\n{projets.Count} projet(s) exporté(s).",
                "Export Excel",
                MessageBoxButton.OK,
                MessageBoxImage.Information);
        }
    }
}
