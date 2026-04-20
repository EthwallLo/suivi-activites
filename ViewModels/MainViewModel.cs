using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Collections.Specialized;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Text.Json;
using MonTableurApp.Models;

namespace MonTableurApp.ViewModels
{
    public class MainViewModel : INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler? PropertyChanged;

        public ObservableCollection<Projet> Projets { get; }

        public List<string> Clients { get; }
        public List<string> Demandeurs { get; }
        public List<string> TypesActivite { get; }
        public List<string> Statuts { get; }

        public int TotalProjets => Projets.Count;
        public int ProjetsEnCours => CountByStatut("En cours");
        public int ProjetsEnAttente => CountByStatut("En attente");
        public int ProjetsTermines => CountByStatut("Termine");

        public MainViewModel()
        {
            Clients = new List<string> { "Client A", "Client B", "Client C" };
            Demandeurs = new List<string> { "Alice", "Bob", "Charlie" };
            TypesActivite = new List<string> { "Analyse", "Dev", "Test" };
            Statuts = new List<string> { "En cours", "Termine", "En attente" };

            Projets = ChargerProjets();
            Projets.CollectionChanged += Projets_CollectionChanged;

            foreach (Projet projet in Projets)
            {
                AttacherProjet(projet);
            }

            NotifierResume();
        }

        private ObservableCollection<Projet> ChargerProjets()
        {
            if (!File.Exists("data.json"))
            {
                return new ObservableCollection<Projet>();
            }

            string json = File.ReadAllText("data.json");
            return JsonSerializer.Deserialize<ObservableCollection<Projet>>(json) ?? new ObservableCollection<Projet>();
        }

        private void Projets_CollectionChanged(object? sender, NotifyCollectionChangedEventArgs e)
        {
            if (e.NewItems != null)
            {
                foreach (Projet projet in e.NewItems)
                {
                    AttacherProjet(projet);
                }
            }

            if (e.OldItems != null)
            {
                foreach (Projet projet in e.OldItems)
                {
                    DetacherProjet(projet);
                }
            }

            Sauvegarder();
            NotifierResume();
        }

        private void AttacherProjet(Projet projet)
        {
            projet.PropertyChanged += Projet_PropertyChanged;
        }

        private void DetacherProjet(Projet projet)
        {
            projet.PropertyChanged -= Projet_PropertyChanged;
        }

        private void Projet_PropertyChanged(object? sender, PropertyChangedEventArgs e)
        {
            Sauvegarder();
            NotifierResume();
        }

        private int CountByStatut(string expectedStatus)
        {
            string expected = NormalizeStatus(expectedStatus);
            return Projets.Count(projet => NormalizeStatus(projet.Statut) == expected);
        }

        private static string NormalizeStatus(string? value)
        {
            return (value ?? string.Empty)
                .Trim()
                .ToLowerInvariant()
                .Replace("é", "e")
                .Replace("è", "e")
                .Replace("ê", "e")
                .Replace("à", "a");
        }

        private void Sauvegarder()
        {
            string json = JsonSerializer.Serialize(Projets, new JsonSerializerOptions { WriteIndented = true });
            File.WriteAllText("data.json", json);
        }

        private void NotifierResume()
        {
            OnPropertyChanged(nameof(TotalProjets));
            OnPropertyChanged(nameof(ProjetsEnCours));
            OnPropertyChanged(nameof(ProjetsEnAttente));
            OnPropertyChanged(nameof(ProjetsTermines));
        }

        private void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}
