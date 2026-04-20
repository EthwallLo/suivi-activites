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
        private const string EssaiFilterAll = "all";
        private const string EssaiFilterToProcess = "to-process";
        private const string EssaiFilterInProgress = "in-progress";
        private const string EssaiFilterDone = "done";

        private static readonly Dictionary<string, List<string>> StatutsParEssai = new(StringComparer.OrdinalIgnoreCase)
        {
            ["Cyclage thermique"] = new List<string> { "\u00C0 faire", "\u00C9chantillon pr\u00EAt", "En cours", "\u00C0 traiter", "Fait", "Non concern\u00E9" },
            ["Traction 100m"] = new List<string> { "\u00C0 faire", "En cours", "Fait", "Non concern\u00E9" },
            ["OTDR"] = new List<string> { "\u00C0 faire", "En cours", "Fait", "Non concern\u00E9" },
            ["Statique Bending"] = new List<string> { "\u00C0 faire", "Fait", "Non concern\u00E9" },
            ["Vieillissement"] = new List<string> { "\u00C0 faire", "En cours", "Fait", "Non concern\u00E9" },
            ["Dimensionnel"] = new List<string> { "\u00C0 faire", "Fait", "Non concern\u00E9" },
            ["Crush"] = new List<string> { "\u00C0 faire", "En cours", "Photos à faire", "Fait", "Non concern\u00E9" },
            ["Cut through"] = new List<string> { "\u00C0 faire", "En cours", "Photos à faire", "Fait", "Non concern\u00E9" },
            ["Kink"] = new List<string> { "\u00C0 faire", "Fait", "Non concern\u00E9" },
            ["Repeated bending"] = new List<string> { "\u00C0 faire", "En cours", "Photos à faire", "Fait", "Non concern\u00E9" },
            ["Torsion"] = new List<string> { "\u00C0 faire", "Fait", "Non concern\u00E9" },
            ["Abrasion marquage"] = new List<string> { "\u00C0 faire", "Fait", "Non concern\u00E9" },
            ["Abrasion gaine"] = new List<string> { "\u00C0 faire", "Photos à faire", "Fait", "Non concern\u00E9" },
            ["Friction gaine"] = new List<string> { "\u00C0 faire", "\u00C9chantillon pr\u00EAt", "En cours", "Fait", "Non concern\u00E9" },
            ["Traction pince"] = new List<string> { "\u00C0 faire", "\u00C9chantillon pr\u00EAt", "En cours", "Fait", "Non concern\u00E9" },
            ["Traction spiral\u00E9"] = new List<string> { "\u00C0 faire", "\u00C9chantillon pr\u00EAt", "En cours", "Fait", "Non concern\u00E9" },
            ["P\u00E9n\u00E9tration d'eau"] = new List<string> { "\u00C0 faire", "\u00C9chantillon pr\u00EAt", "En cours", "Fait", "Non concern\u00E9" },
            ["Petite flamme"] = new List<string> { "\u00C0 faire", "\u00C9chantillon pr\u00EAt", "En cours", "Fait", "Non concern\u00E9" },
            ["Vibration \u00E9olienne"] = new List<string> { "\u00C0 faire", "\u00C9chantillon pr\u00EAt", "En cours", "Fait", "Non concern\u00E9" },
            ["Collage"] = new List<string> { "\u00C0 faire", "Fait", "Non concern\u00E9" },
        };

        private string? searchNomProduit;
        private string? searchNomProduitEssais;
        private string selectedProjetSearchFieldKey = "NomProduit";
        private string activeQuickFilter = QuickFilterAll;
        private string activeEssaiFilter = EssaiFilterAll;
        private int totalProjets;
        private int statutsEnCours;
        private int rapportsEnCours;
        private int projetsFaits;
        private Projet? selectedProjetEssais;
        private int essaisSelectionATraiter;
        private int essaisSelectionEnCours;
        private int essaisSelectionTermines;
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
        public List<SearchFieldOption> ProjetSearchFields { get; }

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

        public string SelectedProjetSearchFieldKey
        {
            get => selectedProjetSearchFieldKey;
            set
            {
                if (selectedProjetSearchFieldKey == value)
                {
                    return;
                }

                selectedProjetSearchFieldKey = value;
                ProjetsView.Refresh();
                OnPropertyChanged(nameof(SelectedProjetSearchFieldKey));
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
        public bool IsEssaiToProcessFilterActive => activeEssaiFilter == EssaiFilterToProcess;
        public bool IsEssaiInProgressFilterActive => activeEssaiFilter == EssaiFilterInProgress;
        public bool IsEssaiDoneFilterActive => activeEssaiFilter == EssaiFilterDone;

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
                RefreshEssaiCollections();
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

        public int EssaisSelectionTermines
        {
            get => essaisSelectionTermines;
            private set
            {
                if (essaisSelectionTermines == value)
                {
                    return;
                }

                essaisSelectionTermines = value;
                OnPropertyChanged(nameof(EssaisSelectionTermines));
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

        public IEnumerable<EssaiSuivi> EssaisPreQualificationFiltres =>
            GetFilteredEssais(SelectedProjetEssais?.EssaisPreQualification ?? Enumerable.Empty<EssaiSuivi>());

        public IEnumerable<EssaiSuivi> EssaisQualificationFiltres =>
            GetFilteredEssais(SelectedProjetEssais?.EssaisQualification ?? Enumerable.Empty<EssaiSuivi>());

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
            ProjetSearchFields = new List<SearchFieldOption>
            {
                new("NomProduit", "Produit"),
                new("NumeroProjet", "Projet")
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
                EssaisSelectionTermines = 0;
                RefreshEssaiCollections();
                return;
            }

            IEnumerable<EssaiSuivi> essaisConcernes = SelectedProjetEssais.Essais.Where(essai => essai.EstConcerne);
            EssaisSelectionTotal = essaisConcernes.Count();
            EssaisSelectionATraiter = essaisConcernes.Count(IsEssaiToProcess);
            EssaisSelectionEnCours = essaisConcernes.Count(IsEssaiInProgress);
            EssaisSelectionTermines = essaisConcernes.Count(IsEssaiDone);
            RefreshEssaiCollections();
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

        public void SetEssaiFilterToToProcess()
        {
            SetEssaiFilter(EssaiFilterToProcess);
        }

        public void SetEssaiFilterToInProgress()
        {
            SetEssaiFilter(EssaiFilterInProgress);
        }

        public void SetEssaiFilterToDone()
        {
            SetEssaiFilter(EssaiFilterDone);
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

        private void SetEssaiFilter(string filter)
        {
            if (activeEssaiFilter == filter)
            {
                activeEssaiFilter = EssaiFilterAll;
            }
            else
            {
                activeEssaiFilter = filter;
            }

            RefreshEssaiCollections();
        }

        private bool MatchesSearch(Projet projet)
        {
            if (string.IsNullOrWhiteSpace(SearchNomProduit))
            {
                return true;
            }

            string searchValue = NormalizeText(SearchNomProduit);

            return SelectedProjetSearchFieldKey switch
            {
                "NumeroProjet" => NormalizeText(projet.NumeroProjet).Contains(searchValue),
                "Client" => NormalizeText(projet.Client).Contains(searchValue),
                "Demandeur" => NormalizeText(projet.Demandeur).Contains(searchValue),
                "TypeActivite" => NormalizeText(projet.TypeActivite).Contains(searchValue),
                "Statut" => NormalizeText(projet.Statut).Contains(searchValue),
                "DossierRacine" => NormalizeText(projet.DossierRacine).Contains(searchValue),
                "Commentaires" => NormalizeText(projet.Commentaires).Contains(searchValue),
                _ => NormalizeText(projet.NomProduit).Contains(searchValue)
            };
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
                statuts = new List<string> { "\u00C0 faire", "\u00C9chantillon pr\u00EAt", "En cours", "\u00C0 traiter", "Fait", "Non concern\u00E9" };
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

        private IEnumerable<EssaiSuivi> GetFilteredEssais(IEnumerable<EssaiSuivi> essais)
        {
            return essais.Where(MatchesEssaiFilter).ToList();
        }

        private bool MatchesEssaiFilter(EssaiSuivi essai)
        {
            return activeEssaiFilter switch
            {
                EssaiFilterToProcess => IsEssaiToProcess(essai),
                EssaiFilterInProgress => IsEssaiInProgress(essai),
                EssaiFilterDone => IsEssaiDone(essai),
                _ => true
            };
        }

        private static bool IsEssaiToProcess(EssaiSuivi essai)
        {
            string statut = NormalizeText(essai.Statut);
            return statut.Contains("trait") && statut != "traite";
        }

        private static bool IsEssaiDone(EssaiSuivi essai)
        {
            return essai.EstConcerne && essai.ProgressionPourcentage == 100;
        }

        private static bool IsEssaiInProgress(EssaiSuivi essai)
        {
            if (!essai.EstConcerne || IsEssaiDone(essai) || IsEssaiToProcess(essai))
            {
                return false;
            }

            string statut = NormalizeText(essai.Statut);
            return !string.IsNullOrWhiteSpace(statut) && statut != "a faire";
        }

        private void RefreshEssaiCollections()
        {
            OnPropertyChanged(nameof(EssaisPreQualificationFiltres));
            OnPropertyChanged(nameof(EssaisQualificationFiltres));
            OnPropertyChanged(nameof(IsEssaiToProcessFilterActive));
            OnPropertyChanged(nameof(IsEssaiInProgressFilterActive));
            OnPropertyChanged(nameof(IsEssaiDoneFilterActive));
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

        public sealed class SearchFieldOption
        {
            public SearchFieldOption(string key, string label)
            {
                Key = key;
                Label = label;
            }

            public string Key { get; }

            public string Label { get; }
        }
    }
}
