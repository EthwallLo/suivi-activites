using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
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

        private void FamillesListBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            FamilleValueTextBox.Text = FamillesListBox.SelectedItem as string ?? string.Empty;
        }

        private void EssaisListBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (EssaisListBox.SelectedItem is not MainViewModel.EssaiDefinitionItem essai)
            {
                return;
            }

            EssaiNameTextBox.Text = essai.Nom;
            EssaiCategoryComboBox.SelectedItem = essai.Categorie;
            EssaiDurationTextBox.Text = essai.DureeHeures.ToString("0.##", CultureInfo.GetCultureInfo("fr-FR"));
            EssaiStatusesTextBox.Text = string.Join(Environment.NewLine, essai.Statuts);
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

        private void AddFamille_Click(object sender, RoutedEventArgs e)
        {
            if (ViewModel is not MainViewModel viewModel)
            {
                return;
            }

            if (viewModel.AddFamilleProduit(FamilleValueTextBox.Text, out string value, out string message))
            {
                FamillesListBox.SelectedItem = value;
                FamilleValueTextBox.Text = string.Empty;
            }

            SetStatus(message);
        }

        private void RenameFamille_Click(object sender, RoutedEventArgs e)
        {
            if (ViewModel is not MainViewModel viewModel)
            {
                return;
            }

            string? selectedValue = FamillesListBox.SelectedItem as string;

            if (viewModel.RenameFamilleProduit(selectedValue, FamilleValueTextBox.Text, out string value, out string message))
            {
                FamillesListBox.SelectedItem = value;
            }

            SetStatus(message);
        }

        private void DeleteFamille_Click(object sender, RoutedEventArgs e)
        {
            if (ViewModel is not MainViewModel viewModel)
            {
                return;
            }

            string? selectedValue = FamillesListBox.SelectedItem as string;
            if (!ConfirmDelete("famille de câbles", selectedValue, viewModel.CountProjectsWithFamilleProduit(selectedValue)))
            {
                return;
            }

            if (viewModel.DeleteFamilleProduit(selectedValue, out string message))
            {
                FamilleValueTextBox.Text = string.Empty;
            }

            SetStatus(message);
        }

        private void AddEssai_Click(object sender, RoutedEventArgs e)
        {
            if (ViewModel is not MainViewModel viewModel || !TryReadEssaiForm(out double durationHours, out List<string> statuses))
            {
                return;
            }

            if (viewModel.AddEssaiDefinition(
                    EssaiNameTextBox.Text,
                    EssaiCategoryComboBox.SelectedItem as string,
                    durationHours,
                    statuses,
                    out MainViewModel.EssaiDefinitionItem? addedItem,
                    out string message))
            {
                EssaisListBox.SelectedItem = addedItem;
            }

            SetStatus(message);
        }

        private void RenameEssai_Click(object sender, RoutedEventArgs e)
        {
            if (ViewModel is not MainViewModel viewModel || !TryReadEssaiForm(out double durationHours, out List<string> statuses))
            {
                return;
            }

            if (viewModel.RenameEssaiDefinition(
                    EssaisListBox.SelectedItem as MainViewModel.EssaiDefinitionItem,
                    EssaiNameTextBox.Text,
                    EssaiCategoryComboBox.SelectedItem as string,
                    durationHours,
                    statuses,
                    out MainViewModel.EssaiDefinitionItem? updatedItem,
                    out string message))
            {
                EssaisListBox.SelectedItem = updatedItem;
            }

            SetStatus(message);
        }

        private void DeleteEssai_Click(object sender, RoutedEventArgs e)
        {
            if (ViewModel is not MainViewModel viewModel)
            {
                return;
            }

            var selectedEssai = EssaisListBox.SelectedItem as MainViewModel.EssaiDefinitionItem;
            if (!ConfirmDelete("essai", selectedEssai?.Nom, viewModel.CountProjectsWithEssai(selectedEssai?.Nom)))
            {
                return;
            }

            if (viewModel.DeleteEssaiDefinition(selectedEssai, out string message))
            {
                EssaiNameTextBox.Text = string.Empty;
                EssaiDurationTextBox.Text = string.Empty;
                EssaiStatusesTextBox.Text = string.Empty;
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

        private bool TryReadEssaiForm(out double durationHours, out List<string> statuses)
        {
            durationHours = 0;
            statuses = EssaiStatusesTextBox.Text
                .Split(new[] { "\r\n", "\n", ";", "," }, StringSplitOptions.RemoveEmptyEntries)
                .Select(status => status.Trim())
                .Where(status => !string.IsNullOrWhiteSpace(status))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();

            string rawDuration = (EssaiDurationTextBox.Text ?? string.Empty).Trim().Replace(',', '.');
            if (!double.TryParse(rawDuration, NumberStyles.AllowDecimalPoint, CultureInfo.InvariantCulture, out durationHours))
            {
                SetStatus("Indique une durée valide en heures.");
                return false;
            }

            if (statuses.Count == 0)
            {
                SetStatus("Ajoute au moins un état disponible.");
                return false;
            }

            return true;
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
