using System;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Media;
using Microsoft.Win32;
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

        private void BrowseDossierRacine_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new OpenFolderDialog
            {
                Title = "Choisir le dossier racine"
            };

            string? currentPath = EditableProjet.DossierRacine;
            if (!string.IsNullOrWhiteSpace(currentPath) && Directory.Exists(currentPath))
            {
                dialog.InitialDirectory = currentPath;
            }

            if (dialog.ShowDialog(this) == true)
            {
                EditableProjet.DossierRacine = dialog.FolderName;
            }
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

        private void ConstrainedCalendarPopup_Opened(object sender, EventArgs e)
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
