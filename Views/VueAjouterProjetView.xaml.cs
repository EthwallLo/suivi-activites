using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using MonTableurApp.Models;
using MonTableurApp.ViewModels;

namespace MonTableurApp.Views
{
    public partial class VueAjouterProjetView : UserControl
    {
        private const string FamilleCable = "Câble";
        private const string FamilleCordon = "Cordon";
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
        private static readonly string[] EssaisCordon =
        {
            "Traction 100m",
            "Statique Bending"
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
                FamilleProduitComboBox.SelectedItem = FamilleCable;
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
                DateDebut = DateDebutTextBox.Text.Trim(),
                DatePrevisionnelle = DatePrevisionnelleTextBox.Text.Trim(),
                DateFin = DateFinTextBox.Text.Trim(),
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
            DateDebutTextBox.Clear();
            DatePrevisionnelleTextBox.Clear();
            DateFinTextBox.Clear();
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
            FamilleProduitComboBox.SelectedItem = viewModel.FamillesProduit.Count > 0 ? FamilleCable : null;
            TypeActiviteComboBox.SelectedItem = viewModel.TypesActivite.Count > 0 ? viewModel.TypesActivite[0] : null;
            StatutComboBox.SelectedItem = viewModel.Statuts.Count > 0 ? viewModel.Statuts[0] : null;
            NumeroProjetTextBox.Focus();
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
            if (famille == FamilleCable)
            {
                SetEssaisSelectionState(_ => true);
            }
            else if (famille == FamilleCordon)
            {
                SetEssaisSelectionState(item => EssaisCordon.Contains(item.NomEssai));
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

        private void SetEssaisSelectionState(System.Func<EssaiSelectionItem, bool> selector)
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
