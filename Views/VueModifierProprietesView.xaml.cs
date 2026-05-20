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

            LoadEssaiForm(essai);
        }

        private void LoadEssaiForm(MainViewModel.EssaiDefinitionItem essai)
        {
            EssaiNameTextBox.Text = essai.Nom;
            EssaiCategoryComboBox.SelectedItem = essai.Categorie;
            EssaiSetupDurationTextBox.Text = essai.DureeMiseEnPlaceHeures.ToString("0.##", CultureInfo.GetCultureInfo("fr-FR"));
            EssaiDurationTextBox.Text = essai.DureeEssaiHeures.ToString("0.##", CultureInfo.GetCultureInfo("fr-FR"));
            EssaiPassageCountTextBox.Text = essai.NombrePassages.ToString(CultureInfo.GetCultureInfo("fr-FR"));
            EssaiRecoveryDurationTextBox.Text = essai.DureeRepriseHeures.ToString("0.##", CultureInfo.GetCultureInfo("fr-FR"));
            EssaiBackgroundCheckBox.IsChecked = essai.EstArrierePlan;
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
            if (ViewModel is not MainViewModel viewModel ||
                !TryReadEssaiForm(out double setupDurationHours, out double testDurationHours, out int passageCount, out double recoveryDurationHours, out bool isBackground, out List<string> statuses))
            {
                return;
            }

            if (viewModel.AddEssaiDefinition(
                    EssaiNameTextBox.Text,
                    EssaiCategoryComboBox.SelectedItem as string,
                    setupDurationHours,
                    testDurationHours,
                    passageCount,
                    recoveryDurationHours,
                    isBackground,
                    statuses,
                    out MainViewModel.EssaiDefinitionItem? addedItem,
                    out string message))
            {
                EssaisListBox.SelectedItem = addedItem;
                if (addedItem != null)
                {
                    LoadEssaiForm(addedItem);
                }
            }

            SetStatus(message);
        }

        private void SaveEssai_Click(object sender, RoutedEventArgs e)
        {
            if (ViewModel is not MainViewModel viewModel ||
                !TryReadEssaiForm(out double setupDurationHours, out double testDurationHours, out int passageCount, out double recoveryDurationHours, out bool isBackground, out List<string> statuses))
            {
                return;
            }

            if (viewModel.UpdateEssaiDefinition(
                    EssaisListBox.SelectedItem as MainViewModel.EssaiDefinitionItem,
                    EssaiNameTextBox.Text,
                    EssaiCategoryComboBox.SelectedItem as string,
                    setupDurationHours,
                    testDurationHours,
                    passageCount,
                    recoveryDurationHours,
                    isBackground,
                    statuses,
                    out MainViewModel.EssaiDefinitionItem? updatedItem,
                    out string message))
            {
                if (updatedItem != null && !ReferenceEquals(EssaisListBox.SelectedItem, updatedItem))
                {
                    EssaisListBox.SelectedItem = updatedItem;
                }

                if (updatedItem != null)
                {
                    LoadEssaiForm(updatedItem);
                }
            }

            SetStatus(message);
        }

        private void RenameEssai_Click(object sender, RoutedEventArgs e)
        {
            if (ViewModel is not MainViewModel viewModel)
            {
                return;
            }

            if (viewModel.RenameEssaiDefinition(
                    EssaisListBox.SelectedItem as MainViewModel.EssaiDefinitionItem,
                    EssaiNameTextBox.Text,
                    out MainViewModel.EssaiDefinitionItem? updatedItem,
                    out string message))
            {
                if (updatedItem != null && !ReferenceEquals(EssaisListBox.SelectedItem, updatedItem))
                {
                    EssaisListBox.SelectedItem = updatedItem;
                }

                if (updatedItem != null)
                {
                    LoadEssaiForm(updatedItem);
                }
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
                EssaiCategoryComboBox.SelectedItem = null;
                EssaiSetupDurationTextBox.Text = string.Empty;
                EssaiDurationTextBox.Text = string.Empty;
                EssaiPassageCountTextBox.Text = string.Empty;
                EssaiRecoveryDurationTextBox.Text = string.Empty;
                EssaiBackgroundCheckBox.IsChecked = false;
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

        private bool TryReadEssaiForm(
            out double setupDurationHours,
            out double testDurationHours,
            out int passageCount,
            out double recoveryDurationHours,
            out bool isBackground,
            out List<string> statuses)
        {
            setupDurationHours = 0;
            testDurationHours = 0;
            passageCount = 1;
            recoveryDurationHours = 0;
            isBackground = EssaiBackgroundCheckBox.IsChecked == true;
            statuses = EssaiStatusesTextBox.Text
                .Split(new[] { "\r\n", "\n", ";", "," }, StringSplitOptions.RemoveEmptyEntries)
                .Select(status => status.Trim())
                .Where(status => !string.IsNullOrWhiteSpace(status))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();

            string rawSetupDuration = (EssaiSetupDurationTextBox.Text ?? string.Empty).Trim().Replace(',', '.');
            string rawDuration = (EssaiDurationTextBox.Text ?? string.Empty).Trim().Replace(',', '.');
            string rawRecoveryDuration = (EssaiRecoveryDurationTextBox.Text ?? string.Empty).Trim().Replace(',', '.');
            string rawPassageCount = (EssaiPassageCountTextBox.Text ?? string.Empty).Trim();
            if (string.IsNullOrWhiteSpace(rawRecoveryDuration))
            {
                rawRecoveryDuration = "0";
            }

            if (string.IsNullOrWhiteSpace(rawPassageCount))
            {
                rawPassageCount = "1";
            }

            if (!double.TryParse(rawSetupDuration, NumberStyles.AllowDecimalPoint, CultureInfo.InvariantCulture, out setupDurationHours) ||
                !double.TryParse(rawDuration, NumberStyles.AllowDecimalPoint, CultureInfo.InvariantCulture, out testDurationHours) ||
                !double.TryParse(rawRecoveryDuration, NumberStyles.AllowDecimalPoint, CultureInfo.InvariantCulture, out recoveryDurationHours) ||
                !int.TryParse(rawPassageCount, NumberStyles.Integer, CultureInfo.InvariantCulture, out passageCount))
            {
                SetStatus("Indique des durees et un nombre de repetitions valides.");
                return false;
            }

            if (setupDurationHours < 0 || testDurationHours < 0 || recoveryDurationHours < 0)
            {
                SetStatus("Les durees ne peuvent pas etre negatives.");
                return false;
            }

            if (passageCount < 1)
            {
                SetStatus("Le nombre de repetitions doit etre au moins egal a 1.");
                return false;
            }

            if (passageCount > 5)
            {
                SetStatus("Le nombre de repetitions ne peut pas depasser 5.");
                return false;
            }

            if (setupDurationHours <= 0 && testDurationHours <= 0 && recoveryDurationHours <= 0)
            {
                SetStatus("Indique au moins une duree superieure a 0.");
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
            StatusBorder.Visibility = Visibility.Visible;
        }
    }
}
