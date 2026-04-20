using System;
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
        private const string QuickFilterAll = "all";
        private const string QuickFilterInProgress = "in-progress";
        private const string QuickFilterReports = "reports";
        private const string QuickFilterDone = "done";

        private static readonly Dictionary<string, List<string>> StatutsParEssai = new(StringComparer.OrdinalIgnoreCase)
        {
            ["Cyclage thermique"] = new List<string> { "\u00C0 faire", "\u00C9chantillon pr\u00EAt", "En cours", "\u00C0 traiter", "Non concern\u00E9" },
            ["Traction 100m"] = new List<string> { "\u00C0 faire", "\u00C9chantillon pr\u00EAt", "En cours", "R\u00E9sultats \u00E0 traiter", "Non concern\u00E9" },
            ["OTDR"] = new List<string> { "\u00C0 faire", "\u00C9chantillon pr\u00EAt", "Mesures en cours", "Courbes \u00E0 traiter", "Non concern\u00E9" },
            ["Statique Bending"] = new List<string> { "\u00C0 faire", "\u00C9chantillon pr\u00EAt", "En cours", "R\u00E9sultats \u00E0 traiter", "Non concern\u00E9" },
            ["Vieillissement"] = new List<string> { "\u00C0 faire", "\u00C9chantillon pr\u00EAt", "En vieillissement", "\u00C0 traiter", "Non concern\u00E9" },
            ["Dimensionnel"] = new List<string> { "\u00C0 faire", "\u00C9chantillon pr\u00EAt", "Mesures en cours", "Mesures \u00E0 valider", "Non concern\u00E9" },
            ["Crush"] = new List<string> { "\u00C0 faire", "\u00C9chantillon pr\u00EAt", "En cours", "R\u00E9sultats \u00E0 traiter", "Non concern\u00E9" },
            ["Cut through"] = new List<string> { "\u00C0 faire", "\u00C9chantillon pr\u00EAt", "En cours", "R\u00E9sultats \u00E0 traiter", "Non concern\u00E9" },
            ["Kink"] = new List<string> { "\u00C0 faire", "\u00C9chantillon pr\u00EAt", "En cours", "R\u00E9sultats \u00E0 traiter", "Non concern\u00E9" },
            ["Repeated bending"] = new List<string> { "\u00C0 faire", "\u00C9chantillon pr\u00EAt", "En cours", "R\u00E9sultats \u00E0 traiter", "Non concern\u00E9" },
            ["Torsion"] = new List<string> { "\u00C0 faire", "\u00C9chantillon pr\u00EAt", "En cours", "R\u00E9sultats \u00E0 traiter", "Non concern\u00E9" },
            ["Abrasion marquage"] = new List<string> { "\u00C0 faire", "\u00C9chantillon pr\u00EAt", "En cours", "\u00C0 contr\u00F4ler", "Non concern\u00E9" },
            ["Abrasion gaine"] = new List<string> { "\u00C0 faire", "\u00C9chantillon pr\u00EAt", "En cours", "\u00C0 contr\u00F4ler", "Non concern\u00E9" },
            ["Friction gaine"] = new List<string> { "\u00C0 faire", "\u00C9chantillon pr\u00EAt", "En cours", "\u00C0 contr\u00F4ler", "Non concern\u00E9" },
            ["Traction pince"] = new List<string> { "\u00C0 faire", "\u00C9chantillon pr\u00EAt", "En cours", "R\u00E9sultats \u00E0 traiter", "Non concern\u00E9" },
            ["Traction spiral\u00E9"] = new List<string> { "\u00C0 faire", "\u00C9chantillon pr\u00EAt", "En cours", "R\u00E9sultats \u00E0 traiter", "Non concern\u00E9" },
            ["P\u00E9n\u00E9tration d'eau"] = new List<string> { "\u00C0 faire", "\u00C9chantillon pr\u00EAt", "En cours", "\u00C0 observer", "Non concern\u00E9" },
            ["Petite flamme"] = new List<string> { "\u00C0 faire", "\u00C9chantillon pr\u00EAt", "En cours", "\u00C0 analyser", "Non concern\u00E9" },
            ["Vibration \u00E9olienne"] = new List<string> { "\u00C0 faire", "\u00C9chantillon pr\u00EAt", "En cours", "\u00C0 analyser", "Non concern\u00E9" },
            ["Collage"] = new List<string> { "\u00C0 faire", "\u00C9chantillon pr\u00EAt", "En cours", "\u00C0 contr\u00F4ler", "Non concern\u00E9" }
        };

        private string? searchNomProduit;
        private string? searchNomProduitEssais;
        private string activeQuickFilter = QuickFilterAll;
        private int totalProjets;
        private int statutsEnCours;
        private int rapportsEnCours;
        private int projetsFaits;
        private Projet? selectedProjetEssais;
        private int essaisSelectionATraiter;
        private int essaisSelectionEnCours;
        private int essaisSelectionTotal;

        public event PropertyChangedEventHandler? PropertyChanged;

        public ObservableCollection<Projet> Projets { get; }
        public ICollectionView ProjetsView { get; }
        public ICollectionView ProjetsEssaisView { get; }

        public List<string> Clients { get; }
        public List<string> Demandeurs { get; }
        public List<string> TypesActivite { get; }
        public List<string> Statuts { get; }
        public List<string> NomsEssais { get; }

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

        public string? SearchNomProduitEssais
        {
            get => searchNomProduitEssais;
            set
            {
                if (searchNomProduitEssais == value)
                {
                    return;
                }

                searchNomProduitEssais = value;
                ProjetsEssaisView.Refresh();
                OnPropertyChanged(nameof(SearchNomProduitEssais));
                EnsureSelectedProjetEssais();
            }
        }

        public bool IsAllFilterActive => activeQuickFilter == QuickFilterAll;
        public bool IsInProgressFilterActive => activeQuickFilter == QuickFilterInProgress;
        public bool IsReportsFilterActive => activeQuickFilter == QuickFilterReports;
        public bool IsDoneFilterActive => activeQuickFilter == QuickFilterDone;

        public Projet? SelectedProjetEssais
        {
            get => selectedProjetEssais;
            set
            {
                if (selectedProjetEssais == value)
                {
                    return;
                }

                selectedProjetEssais = value;
                OnPropertyChanged(nameof(SelectedProjetEssais));
                RefreshSelectedProjectStatistics();
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

        public int EssaisSelectionATraiter
        {
            get => essaisSelectionATraiter;
            private set
            {
                if (essaisSelectionATraiter == value)
                {
                    return;
                }

                essaisSelectionATraiter = value;
                OnPropertyChanged(nameof(EssaisSelectionATraiter));
            }
        }

        public int EssaisSelectionEnCours
        {
            get => essaisSelectionEnCours;
            private set
            {
                if (essaisSelectionEnCours == value)
                {
                    return;
                }

                essaisSelectionEnCours = value;
                OnPropertyChanged(nameof(EssaisSelectionEnCours));
            }
        }

        public int EssaisSelectionTotal
        {
            get => essaisSelectionTotal;
            private set
            {
                if (essaisSelectionTotal == value)
                {
                    return;
                }

                essaisSelectionTotal = value;
                OnPropertyChanged(nameof(EssaisSelectionTotal));
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
            NomsEssais = new List<string>
            {
                "Cyclage thermique",
                "Traction 100m",
                "OTDR",
                "Statique Bending",
                "Vieillissement",
                "Dimensionnel",
                "Crush",
                "Cut through",
                "Kink",
                "Repeated bending",
                "Torsion",
                "Abrasion marquage",
                "Abrasion gaine",
                "Friction gaine",
                "Traction pince",
                "Traction spiral\u00E9",
                "P\u00E9n\u00E9tration d'eau",
                "Petite flamme",
                "Vibration \u00E9olienne",
                "Collage"
            };

            Projets = ChargerProjets();

            foreach (Projet projet in Projets)
            {
                EnsureEssaisForProject(projet);
            }

            ProjetsView = new CollectionViewSource { Source = Projets }.View;
            ProjetsView.Filter = FilterProjet;

            ProjetsEssaisView = new CollectionViewSource { Source = Projets }.View;
            ProjetsEssaisView.Filter = FilterProjetEssais;

            Projets.CollectionChanged += Projets_CollectionChanged;

            foreach (Projet projet in Projets)
            {
                AttacherProjet(projet);
            }

            RefreshStatistics();
            EnsureSelectedProjetEssais();
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

            if (!MatchesSearch(projet))
            {
                return false;
            }

            return MatchesQuickFilter(projet);
        }

        private bool FilterProjetEssais(object obj)
        {
            if (obj is not Projet projet)
            {
                return false;
            }

            if (string.IsNullOrWhiteSpace(SearchNomProduitEssais))
            {
                return true;
            }

            return NormalizeText(projet.NomProduit).Contains(NormalizeText(SearchNomProduitEssais));
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
            ProjetsEssaisView.Refresh();
            Sauvegarder();
            RefreshStatistics();
            EnsureSelectedProjetEssais();
        }

        private void AttacherProjet(Projet projet)
        {
            EnsureEssaisForProject(projet);
            projet.PropertyChanged += Projet_PropertyChanged;
            projet.Essais.CollectionChanged += ProjetEssais_CollectionChanged;

            foreach (EssaiSuivi essai in projet.Essais)
            {
                AttacherEssai(essai);
            }
        }

        private void DetacherProjet(Projet projet)
        {
            projet.PropertyChanged -= Projet_PropertyChanged;
            projet.Essais.CollectionChanged -= ProjetEssais_CollectionChanged;

            foreach (EssaiSuivi essai in projet.Essais)
            {
                DetacherEssai(essai);
            }
        }

        private void Projet_PropertyChanged(object? sender, PropertyChangedEventArgs e)
        {
            if (e.PropertyName == nameof(Projet.NomProduit) || e.PropertyName == nameof(Projet.Statut))
            {
                ProjetsView.Refresh();
                ProjetsEssaisView.Refresh();
                EnsureSelectedProjetEssais();
            }

            Sauvegarder();
            RefreshStatistics();
            RefreshSelectedProjectStatistics();
        }

        private void AttacherEssai(EssaiSuivi essai)
        {
            essai.PropertyChanged += Essai_PropertyChanged;
        }

        private void DetacherEssai(EssaiSuivi essai)
        {
            essai.PropertyChanged -= Essai_PropertyChanged;
        }

        private void ProjetEssais_CollectionChanged(object? sender, NotifyCollectionChangedEventArgs e)
        {
            if (e.NewItems != null)
            {
                foreach (EssaiSuivi essai in e.NewItems)
                {
                    AttacherEssai(essai);
                }
            }

            if (e.OldItems != null)
            {
                foreach (EssaiSuivi essai in e.OldItems)
                {
                    DetacherEssai(essai);
                }
            }

            Sauvegarder();
            RefreshSelectedProjectStatistics();
        }

        private void Essai_PropertyChanged(object? sender, PropertyChangedEventArgs e)
        {
            Sauvegarder();
            RefreshSelectedProjectStatistics();
        }

        private IEnumerable<Projet> GetSearchScopedProjects()
        {
            return Projets.Where(MatchesSearch);
        }

        private void RefreshStatistics()
        {
            TotalProjets = GetSearchScopedProjects().Count();
            StatutsEnCours = CountStatusContaining("cours");
            RapportsEnCours = CountStatusContaining("rapport");
            ProjetsFaits = CountByStatut("fait");
            OnPropertyChanged(nameof(IsAllFilterActive));
            OnPropertyChanged(nameof(IsInProgressFilterActive));
            OnPropertyChanged(nameof(IsReportsFilterActive));
            OnPropertyChanged(nameof(IsDoneFilterActive));
        }

        private void RefreshSelectedProjectStatistics()
        {
            if (SelectedProjetEssais == null)
            {
                EssaisSelectionTotal = 0;
                EssaisSelectionEnCours = 0;
                EssaisSelectionATraiter = 0;
                return;
            }

            EssaisSelectionTotal = SelectedProjetEssais.Essais.Count;
            EssaisSelectionEnCours = SelectedProjetEssais.Essais.Count(essai => NormalizeText(essai.Statut).Contains("cours"));
            EssaisSelectionATraiter = SelectedProjetEssais.Essais.Count(essai => NormalizeText(essai.Statut).Contains("trait"));
        }

        private int CountStatusContaining(string expectedPart)
        {
            string expected = NormalizeText(expectedPart);
            return GetSearchScopedProjects().Count(projet => NormalizeText(projet.Statut).Contains(expected));
        }

        private int CountByStatut(string expectedStatus)
        {
            string expected = NormalizeText(expectedStatus);
            return GetSearchScopedProjects().Count(projet => NormalizeText(projet.Statut) == expected);
        }

        public void SetQuickFilterToAll()
        {
            SetQuickFilter(QuickFilterAll);
        }

        public void SetQuickFilterToInProgress()
        {
            SetQuickFilter(QuickFilterInProgress);
        }

        public void SetQuickFilterToReports()
        {
            SetQuickFilter(QuickFilterReports);
        }

        public void SetQuickFilterToDone()
        {
            SetQuickFilter(QuickFilterDone);
        }

        private void SetQuickFilter(string filter)
        {
            if (activeQuickFilter == filter)
            {
                activeQuickFilter = QuickFilterAll;
            }
            else
            {
                activeQuickFilter = filter;
            }

            ProjetsView.Refresh();
            RefreshStatistics();
        }

        private bool MatchesSearch(Projet projet)
        {
            if (string.IsNullOrWhiteSpace(SearchNomProduit))
            {
                return true;
            }

            return NormalizeText(projet.NomProduit).Contains(NormalizeText(SearchNomProduit));
        }

        private bool MatchesQuickFilter(Projet projet)
        {
            string statut = NormalizeText(projet.Statut);

            return activeQuickFilter switch
            {
                QuickFilterInProgress => statut.Contains("cours"),
                QuickFilterReports => statut.Contains("rapport"),
                QuickFilterDone => statut == "fait",
                _ => true
            };
        }

        private void EnsureEssaisForProject(Projet projet)
        {
            if (projet.Essais == null)
            {
                projet.Essais = new ObservableCollection<EssaiSuivi>();
            }

            foreach (EssaiSuivi essaiExistant in projet.Essais)
            {
                essaiExistant.StatutsDisponibles = CreateStatutsPourEssai(essaiExistant.NomEssai, essaiExistant.Statut);
            }

            foreach (string nomEssai in NomsEssais)
            {
                EssaiSuivi? existingEssai = projet.Essais.FirstOrDefault(
                    essai => string.Equals(essai.NomEssai, nomEssai, StringComparison.OrdinalIgnoreCase));

                if (existingEssai == null)
                {
                    projet.Essais.Add(new EssaiSuivi
                    {
                        NomEssai = nomEssai,
                        Statut = "\u00C0 faire",
                        StatutsDisponibles = CreateStatutsPourEssai(nomEssai, "\u00C0 faire")
                    });
                }
                else
                {
                    existingEssai.StatutsDisponibles = CreateStatutsPourEssai(existingEssai.NomEssai, existingEssai.Statut);
                }
            }
        }

        private static List<string> CreateStatutsPourEssai(string? nomEssai, string? statutActuel)
        {
            string key = nomEssai ?? string.Empty;

            if (!StatutsParEssai.TryGetValue(key, out List<string>? statuts))
            {
                statuts = new List<string> { "\u00C0 faire", "\u00C9chantillon pr\u00EAt", "En cours", "\u00C0 traiter", "Non concern\u00E9" };
            }

            var result = new List<string>(statuts);

            if (!string.IsNullOrWhiteSpace(statutActuel) && !result.Contains(statutActuel))
            {
                result.Add(statutActuel);
            }

            return result;
        }

        private void EnsureSelectedProjetEssais()
        {
            if (SelectedProjetEssais != null && ProjetsEssaisView.Cast<Projet>().Contains(SelectedProjetEssais))
            {
                RefreshSelectedProjectStatistics();
                return;
            }

            SelectedProjetEssais = ProjetsEssaisView.Cast<Projet>().FirstOrDefault();
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
