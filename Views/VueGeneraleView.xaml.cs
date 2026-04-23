using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Threading;
using Microsoft.Win32;
using MonTableurApp.Models;
using MonTableurApp.Services;
using MonTableurApp.ViewModels;

namespace MonTableurApp.Views
{
    public partial class VueGeneraleView : UserControl
    {
        private Window? hostWindow;

        public VueGeneraleView()
        {
            InitializeComponent();
            Loaded += VueGeneraleView_Loaded;
            Unloaded += VueGeneraleView_Unloaded;
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

        private void ProjetsDataGrid_PreviewMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            DependencyObject? source = e.OriginalSource as DependencyObject;
            DataGridCell? cell = FindVisualParent<DataGridCell>(source);

            if (cell is null || cell.IsEditing || !IsSingleClickComboColumn(cell.Column))
            {
                return;
            }

            DataGridRow? row = FindVisualParent<DataGridRow>(cell);
            if (row?.Item is null)
            {
                return;
            }

            ProjetsDataGrid.SelectedItem = row.Item;
            ProjetsDataGrid.CurrentCell = new DataGridCellInfo(row.Item, cell.Column);
            ProjetsDataGrid.BeginEdit();
            e.Handled = true;
        }

        private void VueGeneraleView_Loaded(object sender, RoutedEventArgs e)
        {
            Window? window = Window.GetWindow(this);
            if (window is null || ReferenceEquals(window, hostWindow))
            {
                return;
            }

            if (hostWindow is not null)
            {
                hostWindow.RemoveHandler(PreviewMouseLeftButtonDownEvent, new MouseButtonEventHandler(HostWindow_PreviewMouseLeftButtonDown));
            }

            hostWindow = window;
            hostWindow.AddHandler(
                PreviewMouseLeftButtonDownEvent,
                new MouseButtonEventHandler(HostWindow_PreviewMouseLeftButtonDown),
                true);
        }

        private void VueGeneraleView_Unloaded(object sender, RoutedEventArgs e)
        {
            if (hostWindow is null)
            {
                return;
            }

            hostWindow.RemoveHandler(PreviewMouseLeftButtonDownEvent, new MouseButtonEventHandler(HostWindow_PreviewMouseLeftButtonDown));
            hostWindow = null;
        }

        private void HostWindow_PreviewMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            DependencyObject? source = e.OriginalSource as DependencyObject;

            ComboBox? activeComboBox = FindOpenComboBox(ProjetsDataGrid);
            Popup? clickedPopup = FindVisualParent<Popup>(source);
            ComboBoxItem? clickedComboBoxItem = FindVisualParent<ComboBoxItem>(source);
            ComboBox? clickedComboBox = FindVisualParent<ComboBox>(source) ?? clickedPopup?.PlacementTarget as ComboBox;
            bool clickInsideActiveDropdown =
                activeComboBox is not null
                && (clickedComboBoxItem is not null
                    || ReferenceEquals(clickedPopup?.PlacementTarget, activeComboBox));

            if (clickInsideActiveDropdown)
            {
                return;
            }

            if (activeComboBox is not null)
            {
                activeComboBox.IsDropDownOpen = false;
                ProjetsDataGrid.CommitEdit(DataGridEditingUnit.Cell, true);
                ProjetsDataGrid.CommitEdit(DataGridEditingUnit.Row, true);

                if (clickedComboBox is not null || clickedComboBoxItem is not null)
                {
                    return;
                }
            }
            else if (clickedComboBox is not null || clickedComboBoxItem is not null)
            {
                return;
            }

            DataGrid? clickedGrid = FindVisualParent<DataGrid>(source);
            if (clickedGrid == ProjetsDataGrid)
            {
                DataGridCell? clickedCell = FindVisualParent<DataGridCell>(source);

                if (clickedCell is null)
                {
                    ClearGridSelection();
                }

                return;
            }

            ClearGridSelection();
        }

        private void ProjetsDataGrid_PreparingCellForEdit(object sender, DataGridPreparingCellForEditEventArgs e)
        {
            if (!IsSingleClickComboColumn(e.Column))
            {
                return;
            }

            ComboBox? comboBox = e.EditingElement as ComboBox ?? FindDescendant<ComboBox>(e.EditingElement);
            if (comboBox is null)
            {
                return;
            }

            comboBox.Dispatcher.BeginInvoke(() =>
            {
                comboBox.Focus();
                comboBox.IsDropDownOpen = true;
            }, DispatcherPriority.Background);
        }

        private bool IsSingleClickComboColumn(DataGridColumn? column)
        {
            return column == ClientColumn
                || column == DemandeurColumn
                || column == TypeActiviteColumn
                || column == StatutColumn;
        }

        private static ComboBox? FindOpenComboBox(DependencyObject? parent)
        {
            if (parent is null)
            {
                return null;
            }

            if (parent is ComboBox comboBox && comboBox.IsDropDownOpen)
            {
                return comboBox;
            }

            int childCount = VisualTreeHelper.GetChildrenCount(parent);
            for (int i = 0; i < childCount; i++)
            {
                ComboBox? result = FindOpenComboBox(VisualTreeHelper.GetChild(parent, i));
                if (result is not null)
                {
                    return result;
                }
            }

            return null;
        }

        private void ClearGridSelection()
        {
            ProjetsDataGrid.CommitEdit(DataGridEditingUnit.Cell, true);
            ProjetsDataGrid.CommitEdit(DataGridEditingUnit.Row, true);
            ProjetsDataGrid.CancelEdit(DataGridEditingUnit.Cell);
            ProjetsDataGrid.CancelEdit(DataGridEditingUnit.Row);
            ProjetsDataGrid.UnselectAllCells();
            ProjetsDataGrid.UnselectAll();
            ProjetsDataGrid.SelectedItem = null;
            ProjetsDataGrid.CurrentCell = default;
            RootLayout.Focus();
        }

        private static T? FindVisualParent<T>(DependencyObject? child) where T : DependencyObject
        {
            while (child is not null)
            {
                if (child is T match)
                {
                    return match;
                }

                child = VisualTreeHelper.GetParent(child);
            }

            return null;
        }
    }
}
