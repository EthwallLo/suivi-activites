using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Media;
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

        private void CalendarPopup_SelectedDatesChanged(object sender, SelectionChangedEventArgs e)
        {
            if (sender is not Calendar calendar || e.AddedItems.Count == 0)
            {
                return;
            }

            DateTime? minDate = (calendar.DataContext as Projet)?.DateDebutValue ?? calendar.DisplayDateStart;

            if (minDate is DateTime startDate
                && calendar.SelectedDate is DateTime selectedDate
                && selectedDate.Date < startDate.Date)
            {
                calendar.SelectedDate = startDate.Date;
                return;
            }

            if (calendar.Tag is ToggleButton toggleButton)
            {
                toggleButton.IsChecked = false;
            }
        }

        private void ConstrainedCalendarPopup_Opened(object sender, System.EventArgs e)
        {
            if (sender is not Popup popup)
            {
                return;
            }

            Calendar? calendar = FindDescendant<Calendar>(popup.Child);
            if (popup.DataContext is not Projet projet || calendar is null)
            {
                return;
            }

            calendar.BlackoutDates.Clear();
            calendar.DisplayDateStart = projet.DateDebutValue;

            if (projet.DateDebutValue is DateTime dateDebutValue)
            {
                calendar.BlackoutDates.Add(new CalendarDateRange(DateTime.MinValue, dateDebutValue.Date.AddDays(-1)));
            }
        }

        private static T? FindDescendant<T>(DependencyObject? parent) where T : DependencyObject
        {
            if (parent is null)
            {
                return null;
            }

            if (parent is T match)
            {
                return match;
            }

            int childCount = VisualTreeHelper.GetChildrenCount(parent);
            for (int i = 0; i < childCount; i++)
            {
                T? result = FindDescendant<T>(VisualTreeHelper.GetChild(parent, i));
                if (result is not null)
                {
                    return result;
                }
            }

            return null;
        }
    }
}
