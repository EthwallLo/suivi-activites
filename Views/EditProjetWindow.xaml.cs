using System.Windows;
using MonTableurApp.Models;
using MonTableurApp.ViewModels;

namespace MonTableurApp.Views
{
    public partial class EditProjetWindow : Window
    {
        private readonly Projet sourceProjet;

        public MainViewModel ViewModel { get; }

        public Projet EditableProjet { get; }

        public EditProjetWindow(MainViewModel viewModel, Projet projet)
        {
            InitializeComponent();
            ViewModel = viewModel;
            sourceProjet = projet;
            EditableProjet = CloneProjet(projet);
            DataContext = this;
        }

        private void Save_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrWhiteSpace(EditableProjet.NumeroProjet) ||
                string.IsNullOrWhiteSpace(EditableProjet.NomProduit))
            {
                MessageBox.Show(
                    "Le numéro projet et le nom produit sont obligatoires.",
                    "Champs requis",
                    MessageBoxButton.OK,
                    MessageBoxImage.Warning);
                return;
            }

            if (ViewModel.ProjetExisteAutre(sourceProjet, EditableProjet.NumeroProjet, EditableProjet.NomProduit))
            {
                MessageBox.Show(
                    "Un autre projet avec le même numéro et le même nom produit existe déjà.",
                    "Doublon détecté",
                    MessageBoxButton.OK,
                    MessageBoxImage.Warning);
                return;
            }

            ApplyChanges();
            DialogResult = true;
            Close();
        }

        private void Delete_Click(object sender, RoutedEventArgs e)
        {
            MessageBoxResult result = MessageBox.Show(
                $"Supprimer complètement le produit \"{sourceProjet.NomProduit}\" de l'application ?",
                "Confirmation",
                MessageBoxButton.YesNo,
                MessageBoxImage.Warning);

            if (result != MessageBoxResult.Yes)
            {
                return;
            }

            ViewModel.SupprimerProjet(sourceProjet);
            DialogResult = true;
            Close();
        }

        private void Cancel_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }

        private void ApplyChanges()
        {
            sourceProjet.NumeroProjet = EditableProjet.NumeroProjet;
            sourceProjet.NomProduit = EditableProjet.NomProduit;
            sourceProjet.FamilleProduit = EditableProjet.FamilleProduit;
            sourceProjet.Client = EditableProjet.Client;
            sourceProjet.Demandeur = EditableProjet.Demandeur;
            sourceProjet.TypeActivite = EditableProjet.TypeActivite;
            sourceProjet.DossierRacine = EditableProjet.DossierRacine;
            sourceProjet.Statut = EditableProjet.Statut;
            sourceProjet.DateDebut = EditableProjet.DateDebut;
            sourceProjet.DatePrevisionnelle = EditableProjet.DatePrevisionnelle;
            sourceProjet.DateFin = EditableProjet.DateFin;
            sourceProjet.Commentaires = EditableProjet.Commentaires;
        }

        private static Projet CloneProjet(Projet projet)
        {
            return new Projet
            {
                NumeroProjet = projet.NumeroProjet,
                NomProduit = projet.NomProduit,
                FamilleProduit = projet.FamilleProduit,
                Client = projet.Client,
                Demandeur = projet.Demandeur,
                TypeActivite = projet.TypeActivite,
                DossierRacine = projet.DossierRacine,
                Statut = projet.Statut,
                DateDebut = projet.DateDebut,
                DatePrevisionnelle = projet.DatePrevisionnelle,
                DateFin = projet.DateFin,
                Commentaires = projet.Commentaires
            };
        }
    }
}
