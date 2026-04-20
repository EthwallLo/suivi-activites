using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Collections.Specialized;
using System.ComponentModel;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
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
        public int StatutsEnCours => CountStatusContaining("cours");
        public int RapportsEnCours => CountStatusContaining("rapport");
        public int ProjetsFaits => CountByStatut("fait");

        public MainViewModel()
        {
            Clients = new List<string> { "Orange", "Free", "Bouygues", "DTAG", "BT", "N/A" };
            Demandeurs = new List<string> { "RUC", "JEN", "KYJ", "JLC", "JER", "JEL", "JYM", "XAL" };
            TypesActivite = new List<string> { "Qualification", "Appel d'offre", "Investigation", "Caract\u00E9risation" };
            Statuts = new List<string>
            {
                "\u00C0 faire",
                "Pr\u00E9-qualification en cours",
                "Qualification en cours",
                "Rapport en cours",
                "Rapport termin\u00E9",
                "Fait"
            };

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

        private int CountStatusContaining(string expectedPart)
        {
            string expected = NormalizeText(expectedPart);
            return Projets.Count(projet => NormalizeText(projet.Statut).Contains(expected));
        }

        private int CountByStatut(string expectedStatus)
        {
            string expected = NormalizeText(expectedStatus);
            return Projets.Count(projet => NormalizeText(projet.Statut) == expected);
        }

        private static string NormalizeText(string? value)
        {
            string normalized = (value ?? string.Empty).Trim().ToLowerInvariant().Normalize(NormalizationForm.FormD);
            var builder = new StringBuilder(normalized.Length);

            foreach (char current in normalized)
            {
                if (CharUnicodeInfo.GetUnicodeCategory(current) != UnicodeCategory.NonSpacingMark)
                {
                    builder.Append(current);
                }
            }

            return builder.ToString().Normalize(NormalizationForm.FormC);
        }

        private void Sauvegarder()
        {
            string json = JsonSerializer.Serialize(Projets, new JsonSerializerOptions { WriteIndented = true });
            File.WriteAllText("data.json", json);
        }

        private void NotifierResume()
        {
            OnPropertyChanged(nameof(TotalProjets));
            OnPropertyChanged(nameof(StatutsEnCours));
            OnPropertyChanged(nameof(RapportsEnCours));
            OnPropertyChanged(nameof(ProjetsFaits));
        }

        private void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}
