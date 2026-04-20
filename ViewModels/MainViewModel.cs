using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Collections.Specialized;
using System.ComponentModel;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.Encodings.Web;
using System.Text.Json;
using System.Windows.Data;
using MonTableurApp.Models;

namespace MonTableurApp.ViewModels
{
    public class MainViewModel : INotifyPropertyChanged
    {
        private static readonly string DataFilePath = Path.Combine(AppContext.BaseDirectory, "data.json");

        private string? searchNomProduit;
        private int totalProjets;
        private int statutsEnCours;
        private int rapportsEnCours;
        private int projetsFaits;

        public event PropertyChangedEventHandler? PropertyChanged;

        public ObservableCollection<Projet> Projets { get; }
        public ICollectionView ProjetsView { get; }

        public List<string> Clients { get; }
        public List<string> Demandeurs { get; }
        public List<string> TypesActivite { get; }
        public List<string> Statuts { get; }

        public string? SearchNomProduit
        {
            get => searchNomProduit;
            set
            {
                if (searchNomProduit == value)
                {
                    return;
                }

                searchNomProduit = value;
                ProjetsView.Refresh();
                OnPropertyChanged(nameof(SearchNomProduit));
                RefreshStatistics();
            }
        }

        public int TotalProjets
        {
            get => totalProjets;
            private set
            {
                if (totalProjets == value)
                {
                    return;
                }

                totalProjets = value;
                OnPropertyChanged(nameof(TotalProjets));
            }
        }

        public int StatutsEnCours
        {
            get => statutsEnCours;
            private set
            {
                if (statutsEnCours == value)
                {
                    return;
                }

                statutsEnCours = value;
                OnPropertyChanged(nameof(StatutsEnCours));
            }
        }

        public int RapportsEnCours
        {
            get => rapportsEnCours;
            private set
            {
                if (rapportsEnCours == value)
                {
                    return;
                }

                rapportsEnCours = value;
                OnPropertyChanged(nameof(RapportsEnCours));
            }
        }

        public int ProjetsFaits
        {
            get => projetsFaits;
            private set
            {
                if (projetsFaits == value)
                {
                    return;
                }

                projetsFaits = value;
                OnPropertyChanged(nameof(ProjetsFaits));
            }
        }

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
            ProjetsView = CollectionViewSource.GetDefaultView(Projets);
            ProjetsView.Filter = FilterProjet;
            Projets.CollectionChanged += Projets_CollectionChanged;

            foreach (Projet projet in Projets)
            {
                AttacherProjet(projet);
            }

            RefreshStatistics();
        }

        private ObservableCollection<Projet> ChargerProjets()
        {
            if (!File.Exists(DataFilePath))
            {
                return new ObservableCollection<Projet>();
            }

            string json = File.ReadAllText(DataFilePath, Encoding.UTF8);
            return JsonSerializer.Deserialize<ObservableCollection<Projet>>(json) ?? new ObservableCollection<Projet>();
        }

        private bool FilterProjet(object obj)
        {
            if (obj is not Projet projet)
            {
                return false;
            }

            if (string.IsNullOrWhiteSpace(SearchNomProduit))
            {
                return true;
            }

            return NormalizeText(projet.NomProduit).Contains(NormalizeText(SearchNomProduit));
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

            ProjetsView.Refresh();
            Sauvegarder();
            RefreshStatistics();
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
            if (e.PropertyName == nameof(Projet.NomProduit))
            {
                ProjetsView.Refresh();
            }

            Sauvegarder();
            RefreshStatistics();
        }

        private IEnumerable<Projet> GetVisibleProjects()
        {
            return ProjetsView.Cast<Projet>();
        }

        private void RefreshStatistics()
        {
            TotalProjets = GetVisibleProjects().Count();
            StatutsEnCours = CountStatusContaining("cours");
            RapportsEnCours = CountStatusContaining("rapport");
            ProjetsFaits = CountByStatut("fait");
        }

        private int CountStatusContaining(string expectedPart)
        {
            string expected = NormalizeText(expectedPart);
            return GetVisibleProjects().Count(projet => NormalizeText(projet.Statut).Contains(expected));
        }

        private int CountByStatut(string expectedStatus)
        {
            string expected = NormalizeText(expectedStatus);
            return GetVisibleProjects().Count(projet => NormalizeText(projet.Statut) == expected);
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
            string json = JsonSerializer.Serialize(
                Projets,
                new JsonSerializerOptions
                {
                    WriteIndented = true,
                    Encoder = JavaScriptEncoder.UnsafeRelaxedJsonEscaping
                });

            File.WriteAllText(DataFilePath, json, new UTF8Encoding(false));
        }

        private void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}
