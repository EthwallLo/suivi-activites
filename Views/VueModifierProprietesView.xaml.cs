using System;
using System.Windows;
using System.Windows.Controls;
using MonTableurApp.ViewModels;

namespace MonTableurApp.Views
{
    public partial class VueModifierProprietesView : UserControl
    {
        public VueModifierProprietesView()
        {
            InitializeComponent();
        }

        private MainViewModel? ViewModel => DataContext as MainViewModel;

        private void ClientsListBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            ClientValueTextBox.Text = ClientsListBox.SelectedItem as string ?? string.Empty;
        }

        private void DemandeursListBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            DemandeurValueTextBox.Text = DemandeursListBox.SelectedItem as string ?? string.Empty;
        }

        private void AddClient_Click(object sender, RoutedEventArgs e)
        {
            if (ViewModel is not MainViewModel viewModel)
            {
                return;
            }

            if (viewModel.AddClient(ClientValueTextBox.Text, out string value, out string message))
            {
                ClientsListBox.SelectedItem = value;
                ClientValueTextBox.Text = string.Empty;
            }

            SetStatus(message);
        }

        private void RenameClient_Click(object sender, RoutedEventArgs e)
        {
            if (ViewModel is not MainViewModel viewModel)
            {
                return;
            }

            string? selectedValue = ClientsListBox.SelectedItem as string;

            if (viewModel.RenameClient(selectedValue, ClientValueTextBox.Text, out string value, out string message))
            {
                ClientsListBox.SelectedItem = value;
            }

            SetStatus(message);
        }

        private void DeleteClient_Click(object sender, RoutedEventArgs e)
        {
            if (ViewModel is not MainViewModel viewModel)
            {
                return;
            }

            string? selectedValue = ClientsListBox.SelectedItem as string;
            if (!ConfirmDelete("client", selectedValue, viewModel.CountProjectsWithClient(selectedValue)))
            {
                return;
            }

            if (viewModel.DeleteClient(selectedValue, out string message))
            {
                ClientValueTextBox.Text = string.Empty;
            }

            SetStatus(message);
        }

        private void AddDemandeur_Click(object sender, RoutedEventArgs e)
        {
            if (ViewModel is not MainViewModel viewModel)
            {
                return;
            }

            if (viewModel.AddDemandeur(DemandeurValueTextBox.Text, out string value, out string message))
            {
                DemandeursListBox.SelectedItem = value;
                DemandeurValueTextBox.Text = string.Empty;
            }

            SetStatus(message);
        }

        private void RenameDemandeur_Click(object sender, RoutedEventArgs e)
        {
            if (ViewModel is not MainViewModel viewModel)
            {
                return;
            }

            string? selectedValue = DemandeursListBox.SelectedItem as string;

            if (viewModel.RenameDemandeur(selectedValue, DemandeurValueTextBox.Text, out string value, out string message))
            {
                DemandeursListBox.SelectedItem = value;
            }

            SetStatus(message);
        }

        private void DeleteDemandeur_Click(object sender, RoutedEventArgs e)
        {
            if (ViewModel is not MainViewModel viewModel)
            {
                return;
            }

            string? selectedValue = DemandeursListBox.SelectedItem as string;
            if (!ConfirmDelete("demandeur", selectedValue, viewModel.CountProjectsWithDemandeur(selectedValue)))
            {
                return;
            }

            if (viewModel.DeleteDemandeur(selectedValue, out string message))
            {
                DemandeurValueTextBox.Text = string.Empty;
            }

            SetStatus(message);
        }

        private void SelectAllDefaultEssais_Click(object sender, RoutedEventArgs e)
        {
            if (ViewModel is not MainViewModel viewModel)
            {
                return;
            }

            viewModel.SetAllEssaisParDefautSelection(true);
            SetStatus("Tous les essais sont cochés pour cette famille.");
        }

        private void ClearDefaultEssais_Click(object sender, RoutedEventArgs e)
        {
            if (ViewModel is not MainViewModel viewModel)
            {
                return;
            }

            viewModel.SetAllEssaisParDefautSelection(false);
            SetStatus("Tous les essais sont décochés pour cette famille.");
        }

        private void SaveDefaultEssais_Click(object sender, RoutedEventArgs e)
        {
            if (ViewModel is not MainViewModel viewModel)
            {
                return;
            }

            viewModel.SaveSelectedEssaisParDefaut(out string message);
            SetStatus(message);
        }

        private bool ConfirmDelete(string label, string? value, int usedCount)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                SetStatus($"Sélectionne un {label} à supprimer.");
                return false;
            }

            if (usedCount == 0)
            {
                return true;
            }

            MessageBoxResult result = MessageBox.Show(
                $"{value} est utilisé dans {usedCount} projet(s). Le supprimer retirera cette valeur des projets concernés.",
                "Confirmer la suppression",
                MessageBoxButton.YesNo,
                MessageBoxImage.Warning);

            return result == MessageBoxResult.Yes;
        }

        private void SetStatus(string message)
        {
            StatusTextBlock.Text = string.IsNullOrWhiteSpace(message)
                ? "Action terminée."
                : message;
        }
    }
}
