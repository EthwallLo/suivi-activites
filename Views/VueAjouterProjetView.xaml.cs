using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Globalization;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using Microsoft.Win32;
using MonTableurApp.Models;
using MonTableurApp.ViewModels;

namespace MonTableurApp.Views
{
    public partial class VueAjouterProjetView : UserControl
    {
        private static readonly CultureInfo DateCulture = CultureInfo.GetCultureInfo("fr-FR");
        private const string FamilleCableIndoor = "Câble indoor";
        private const string FamilleCableOutdoor = "Câble outdoor";
        private const string FamilleCableDrop = "Câble drop";
        private const string FamilleCordon = "Cordon";
        private const string FamillePatchcords = "Patchcords";
        private const string FamilleAramides = "Aramides";
        private const string FamilleRipcords = "Ripcords";
        private const string FamilleFrp = "FRP";
        private const string FamilleAutre = "Autre";

        private static readonly string[] EssaisPreQualification =
        {
            "Traction 100m",
            "Cyclage thermique",
            "Statique Bending",
            "Vieillissement"
        };

        private static readonly string[] EssaisExterieurs =
        {
            "Exposition UV",
            "CPR"
        };

        private static readonly string[] TousLesEssais =
        {
            "Traction 100m",
            "Cyclage thermique",
            "OTDR",
            "Statique Bending",
            "Vieillissement",
            "Dimensionnel",
            "Crush",
            "Cut through",
            "Traction verticale",
            "Kink",
            "Repeated bending",
            "Torsion",
            "Abrasion marquage",
            "Abrasion gaine",
            "Friction gaine",
            "Traction pince",
            "Traction spiralé",
            "Pénétration d'eau",
            "Petite flamme",
            "Vibration éolienne",
            "Collage",
            "Exposition UV",
            "CPR"
        };

        private static readonly string[] EssaisCordon =
        {
            "Traction 100m",
            "Statique Bending"
        };

        private static readonly Dictionary<string, HashSet<string>> EssaisPresetsParFamille = new(StringComparer.Ordinal)
        {
            [FamilleCableIndoor] = new HashSet<string>(TousLesEssais, StringComparer.Ordinal),
            [FamilleCableOutdoor] = new HashSet<string>(TousLesEssais, StringComparer.Ordinal),
            [FamilleCableDrop] = new HashSet<string>(TousLesEssais, StringComparer.Ordinal),
            [FamilleCordon] = new HashSet<string>(new[]
            {
                "Traction 100m",
                "OTDR",
                "Statique Bending",
                "Repeated bending",
                "Torsion",
                "Exposition UV"
            }, StringComparer.Ordinal),
            [FamillePatchcords] = new HashSet<string>(new[]
            {
                "Traction 100m",
                "OTDR",
                "Statique Bending",
                "Repeated bending",
                "Cut through",
                "Abrasion gaine",
                "CPR"
            }, StringComparer.Ordinal),
            [FamilleAramides] = new HashSet<string>(new[]
            {
                "Traction verticale"
            }, StringComparer.Ordinal),
            [FamilleRipcords] = new HashSet<string>(new[]
            {
                "Traction verticale"
            }, StringComparer.Ordinal),
            [FamilleFrp] = new HashSet<string>(new[]
            {
                "Traction verticale",
                "Vieillissement"
            }, StringComparer.Ordinal),
            [FamilleAutre] = new HashSet<string>(new[]
            {
                "Cyclage thermique",
                "OTDR",
                "Dimensionnel",
                "Crush",
                "Friction gaine",
                "Pénétration d'eau",
                "Collage"
            }, StringComparer.Ordinal)
        };

        public ObservableCollection<EssaiSelectionItem> EssaisPreQualificationSelection { get; } = new();
        public ObservableCollection<EssaiSelectionItem> EssaisQualificationSelection { get; } = new();
        public ObservableCollection<EssaiSelectionItem> EssaisExterieursSelection { get; } = new();

        public VueAjouterProjetView()
        {
            InitializeComponent();
            Loaded += VueAjouterProjetView_Loaded;
        }

        private void VueAjouterProjetView_Loaded(object sender, RoutedEventArgs e)
        {
            if (DataContext is not MainViewModel viewModel)
            {
                return;
            }

            EnsureEssaiSelections(viewModel);
            UpdateToggleEssaisSelectionButtonText();

            if (FamilleProduitComboBox.SelectedItem == null && viewModel.FamillesProduit.Count > 0)
            {
                FamilleProduitComboBox.SelectedItem = FamilleCableIndoor;
            }

            if (ClientComboBox.SelectedItem == null && viewModel.Clients.Count > 0)
            {
                ClientComboBox.SelectedItem = viewModel.Clients[0];
            }

            if (DemandeurComboBox.SelectedItem == null && viewModel.Demandeurs.Count > 0)
            {
                DemandeurComboBox.SelectedItem = viewModel.Demandeurs[0];
            }

            if (TypeActiviteComboBox.SelectedItem == null && viewModel.TypesActivite.Count > 0)
            {
                TypeActiviteComboBox.SelectedItem = viewModel.TypesActivite[0];
            }

            if (StatutComboBox.SelectedItem == null && viewModel.Statuts.Count > 0)
            {
                StatutComboBox.SelectedItem = viewModel.Statuts[0];
            }
        }

        private void AjouterProjet_Click(object sender, RoutedEventArgs e)
        {
            if (DataContext is not MainViewModel viewModel)
            {
                return;
            }

            string numeroProjet = NumeroProjetTextBox.Text.Trim();
            string nomProduit = NomProduitTextBox.Text.Trim();

            if (string.IsNullOrWhiteSpace(numeroProjet) || string.IsNullOrWhiteSpace(nomProduit))
            {
                MessageBox.Show(
                    "Le numéro projet et le nom produit sont obligatoires.",
                    "Projet incomplet",
                    MessageBoxButton.OK,
                    MessageBoxImage.Warning);
                return;
            }

            if (viewModel.ProjetExiste(numeroProjet, nomProduit))
            {
                MessageBox.Show(
                    "Un projet avec ce numéro et ce nom produit existe déjà.",
                    "Doublon détecté",
                    MessageBoxButton.OK,
                    MessageBoxImage.Information);
                return;
            }

            var projet = new Projet
            {
                NumeroProjet = numeroProjet,
                NomProduit = nomProduit,
                FamilleProduit = FamilleProduitComboBox.SelectedItem as string,
                Client = ClientComboBox.SelectedItem as string,
                Demandeur = DemandeurComboBox.SelectedItem as string,
                TypeActivite = TypeActiviteComboBox.SelectedItem as string,
                DossierRacine = DossierRacineTextBox.Text.Trim(),
                Statut = StatutComboBox.SelectedItem as string,
                DateDebut = FormatDate(DateDebutCalendar.SelectedDate),
                DatePrevisionnelle = FormatDate(DatePrevisionnelleCalendar.SelectedDate),
                DateFin = FormatDate(DateFinCalendar.SelectedDate),
                Commentaires = CommentairesTextBox.Text.Trim(),
                Essais = new ObservableCollection<EssaiSuivi>(BuildEssaisForNewProject())
            };

            viewModel.AjouterProjet(projet);
            ViderFormulaire();

            MessageBox.Show(
                "Le projet a bien été ajouté.",
                "Projet créé",
                MessageBoxButton.OK,
                MessageBoxImage.Information);
        }

        private void ViderFormulaire_Click(object sender, RoutedEventArgs e)
        {
            ViderFormulaire();
        }

        private void ViderFormulaire()
        {
            NumeroProjetTextBox.Clear();
            NomProduitTextBox.Clear();
            DossierRacineTextBox.Clear();
            DateDebutCalendar.SelectedDate = null;
            DatePrevisionnelleCalendar.SelectedDate = null;
            DateFinCalendar.SelectedDate = null;
            CommentairesTextBox.Clear();

            foreach (EssaiSelectionItem item in GetAllEssaisSelectionItems())
            {
                item.IsSelected = true;
            }

            UpdateToggleEssaisSelectionButtonText();

            if (DataContext is not MainViewModel viewModel)
            {
                return;
            }

            ClientComboBox.SelectedItem = viewModel.Clients.Count > 0 ? viewModel.Clients[0] : null;
            DemandeurComboBox.SelectedItem = viewModel.Demandeurs.Count > 0 ? viewModel.Demandeurs[0] : null;
            FamilleProduitComboBox.SelectedItem = viewModel.FamillesProduit.Count > 0 ? FamilleCableIndoor : null;
            TypeActiviteComboBox.SelectedItem = viewModel.TypesActivite.Count > 0 ? viewModel.TypesActivite[0] : null;
            StatutComboBox.SelectedItem = viewModel.Statuts.Count > 0 ? viewModel.Statuts[0] : null;
            NumeroProjetTextBox.Focus();
        }

        private void ParcourirDossierRacine_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new OpenFolderDialog
            {
                Title = "Sélectionner le dossier racine du projet",
                InitialDirectory = string.IsNullOrWhiteSpace(DossierRacineTextBox.Text)
                    ? Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
                    : DossierRacineTextBox.Text.Trim()
            };

            if (dialog.ShowDialog() == true)
            {
                DossierRacineTextBox.Text = dialog.FolderName;
            }
        }

        private void ToggleEssaisSelection_Click(object sender, RoutedEventArgs e)
        {
            bool shouldSelectAll = !AreAllEssaisSelected();

            foreach (EssaiSelectionItem item in GetAllEssaisSelectionItems())
            {
                item.IsSelected = shouldSelectAll;
            }

            UpdateToggleEssaisSelectionButtonText();
        }

        private void FamilleProduit_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (!IsLoaded)
            {
                return;
            }

            string? famille = FamilleProduitComboBox.SelectedItem as string;
            if (!string.IsNullOrWhiteSpace(famille) &&
                     EssaisPresetsParFamille.TryGetValue(famille, out HashSet<string>? essaisPreselectionnes))
            {
                SetEssaisSelectionState(item => essaisPreselectionnes.Contains(item.NomEssai));
            }

            UpdateToggleEssaisSelectionButtonText();
        }

        private void EnsureEssaiSelections(MainViewModel viewModel)
        {
            if (EssaisPreQualificationSelection.Count > 0 ||
                EssaisQualificationSelection.Count > 0 ||
                EssaisExterieursSelection.Count > 0)
            {
                return;
            }

            IEnumerable<string> essaisQualification = viewModel.NomsEssais
                .Where(nomEssai =>
                    !EssaisPreQualification.Contains(nomEssai) &&
                    !EssaisExterieurs.Contains(nomEssai));

            foreach (string nomEssai in EssaisPreQualification)
            {
                AddEssaiSelectionItem(EssaisPreQualificationSelection, nomEssai);
            }

            foreach (string nomEssai in essaisQualification)
            {
                AddEssaiSelectionItem(EssaisQualificationSelection, nomEssai);
            }

            foreach (string nomEssai in EssaisExterieurs)
            {
                AddEssaiSelectionItem(EssaisExterieursSelection, nomEssai);
            }
        }

        private void AddEssaiSelectionItem(ObservableCollection<EssaiSelectionItem> target, string nomEssai)
        {
            var item = new EssaiSelectionItem(nomEssai);
            item.PropertyChanged += EssaiSelectionItem_PropertyChanged;
            target.Add(item);
        }

        private void EssaiSelectionItem_PropertyChanged(object? sender, PropertyChangedEventArgs e)
        {
            if (e.PropertyName == nameof(EssaiSelectionItem.IsSelected))
            {
                UpdateToggleEssaisSelectionButtonText();
            }
        }

        private IEnumerable<EssaiSuivi> BuildEssaisForNewProject()
        {
            return GetAllEssaisSelectionItems()
                .Select(item => new EssaiSuivi
                {
                    NomEssai = item.NomEssai,
                    Statut = item.IsSelected ? "À faire" : "Non concerné"
                })
                .ToList();
        }

        private bool AreAllEssaisSelected()
        {
            return GetAllEssaisSelectionItems().All(item => item.IsSelected);
        }

        private void UpdateToggleEssaisSelectionButtonText()
        {
            if (ToggleEssaisSelectionButton == null)
            {
                return;
            }

            ToggleEssaisSelectionButton.Content = AreAllEssaisSelected()
                ? "Tout désélectionner"
                : "Tout sélectionner";
        }

        private void SetEssaisSelectionState(Func<EssaiSelectionItem, bool> selector)
        {
            foreach (EssaiSelectionItem item in GetAllEssaisSelectionItems())
            {
                item.IsSelected = selector(item);
            }
        }

        private IEnumerable<EssaiSelectionItem> GetAllEssaisSelectionItems()
        {
            return EssaisPreQualificationSelection
                .Concat(EssaisQualificationSelection)
                .Concat(EssaisExterieursSelection);
        }

        private static string FormatDate(DateTime? date)
        {
            return date?.ToString("dd/MM/yyyy", DateCulture) ?? string.Empty;
        }

        public sealed class EssaiSelectionItem : INotifyPropertyChanged
        {
            private bool isSelected = true;

            public EssaiSelectionItem(string nomEssai)
            {
                NomEssai = nomEssai;
            }

            public event PropertyChangedEventHandler? PropertyChanged;

            public string NomEssai { get; }

            public bool IsSelected
            {
                get => isSelected;
                set
                {
                    if (isSelected == value)
                    {
                        return;
                    }

                    isSelected = value;
                    PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(IsSelected)));
                }
            }
        }
    }
}
