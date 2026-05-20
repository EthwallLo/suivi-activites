using System;
using System.Globalization;
using System.Windows;
using MonTableurApp.Models;
using MonTableurApp.ViewModels;

namespace MonTableurApp.Views
{
    public partial class EditEssaiPlanningWindow : Window
    {
        private static readonly CultureInfo FrenchCulture = CultureInfo.GetCultureInfo("fr-FR");
        private readonly MainViewModel viewModel;
        private readonly EssaiSuivi essai;

        public EditEssaiPlanningWindow(MainViewModel viewModel, Projet projet, EssaiSuivi essai)
        {
            InitializeComponent();
            this.viewModel = viewModel;
            this.essai = essai;

            TitleTextBlock.Text = essai.NomEssai ?? "Essai";
            SubtitleTextBlock.Text = $"{projet.NomProduit} · {projet.NumeroProjet}";
            LoadValues();
        }

        private void LoadValues()
        {
            MainViewModel.EssaiPlanningValues defaults = viewModel.GetDefaultEssaiPlanning(essai);
            MainViewModel.EssaiPlanningValues values = viewModel.GetEffectiveEssaiPlanning(essai);

            DefaultSummaryTextBlock.Text =
                $"Défaut : mise {FormatHours(defaults.DureeMiseEnPlaceHeures)} · essai {FormatHours(defaults.DureeEssaiHeures)} · reprise {FormatHours(defaults.DureeRepriseHeures)}" +
                (defaults.EstArrierePlan ? " · fond" : string.Empty);

            SetupDurationTextBox.Text = values.DureeMiseEnPlaceHeures.ToString("0.##", FrenchCulture);
            TestDurationTextBox.Text = values.DureeEssaiHeures.ToString("0.##", FrenchCulture);
            RecoveryDurationTextBox.Text = values.DureeRepriseHeures.ToString("0.##", FrenchCulture);
            BackgroundCheckBox.IsChecked = values.EstArrierePlan;
            StatusTextBlock.Text = essai.HasCustomPlanning
                ? "Ces valeurs remplacent le défaut pour cet essai uniquement."
                : "Cet essai utilise encore les valeurs par défaut.";
        }

        private void Save_Click(object sender, RoutedEventArgs e)
        {
            if (!TryReadHours(SetupDurationTextBox.Text, out double setupHours) ||
                !TryReadHours(TestDurationTextBox.Text, out double testHours) ||
                !TryReadHours(RecoveryDurationTextBox.Text, out double recoveryHours))
            {
                StatusTextBlock.Text = "Indique des durées valides en heures.";
                return;
            }

            if (!viewModel.ApplyEssaiPlanningConfiguration(
                    essai,
                    setupHours,
                    testHours,
                    recoveryHours,
                    BackgroundCheckBox.IsChecked == true,
                    out string message))
            {
                StatusTextBlock.Text = message;
                return;
            }

            DialogResult = true;
            Close();
        }

        private void Reset_Click(object sender, RoutedEventArgs e)
        {
            viewModel.ResetEssaiPlanningConfiguration(essai);
            DialogResult = true;
            Close();
        }

        private void Cancel_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }

        private static bool TryReadHours(string? value, out double hours)
        {
            string normalizedValue = (value ?? string.Empty).Trim().Replace(',', '.');
            return double.TryParse(
                normalizedValue,
                NumberStyles.AllowDecimalPoint,
                CultureInfo.InvariantCulture,
                out hours);
        }

        private static string FormatHours(double hours)
        {
            return $"{hours.ToString("0.##", FrenchCulture)} h";
        }
    }
}
