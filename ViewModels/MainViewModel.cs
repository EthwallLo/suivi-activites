using System;
using System.Collections;
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
using System.Windows;
using System.Windows.Data;
using MonTableurApp.Models;

namespace MonTableurApp.ViewModels
{
    public class MainViewModel : INotifyPropertyChanged
    {
        private static readonly string DataFilePath = Path.Combine(AppContext.BaseDirectory, "data.json");
        private static readonly string ReferencePropertiesPath = Path.Combine(AppContext.BaseDirectory, "reference-properties.json");
        private static readonly string[] DefaultClients = { "Orange", "Free", "Bouygues", "DTAG", "BT", "N/A" };
        private static readonly string[] DefaultDemandeurs = { "RUC", "JEN", "KYJ", "JLC", "JER", "JEL", "JYM", "XAL" };
        private static readonly string[] DefaultFamillesProduit = { "Câble indoor", "Câble outdoor", "Câble drop", "Cordon", "Patchcords", "Aramides", "Ripcords", "FRP", "Autre" };
        private static readonly string[] DefaultTypesActivite = { "Qualification", "Appel d'offre", "Investigation", "Caractérisation" };
        private static readonly string[] DefaultNomsEssais =
        {
            "Cyclage thermique",
            "Traction 100m",
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

        private static readonly Dictionary<string, string[]> DefaultEssaisParFamille = new(StringComparer.OrdinalIgnoreCase)
        {
            ["Câble indoor"] = DefaultNomsEssais,
            ["Câble outdoor"] = DefaultNomsEssais,
            ["Câble drop"] = DefaultNomsEssais,
            ["Cordon"] = new[]
            {
                "Traction 100m",
                "OTDR",
                "Statique Bending",
                "Repeated bending",
                "Torsion",
                "Exposition UV"
            },
            ["Patchcords"] = new[]
            {
                "Traction 100m",
                "OTDR",
                "Statique Bending",
                "Repeated bending",
                "Cut through",
                "Abrasion gaine",
                "CPR"
            },
            ["Aramides"] = new[]
            {
                "Traction verticale"
            },
            ["Ripcords"] = new[]
            {
                "Traction verticale"
            },
            ["FRP"] = new[]
            {
                "Traction verticale",
                "Vieillissement"
            },
            ["Autre"] = new[]
            {
                "Cyclage thermique",
                "OTDR",
                "Dimensionnel",
                "Crush",
                "Friction gaine",
                "Pénétration d'eau",
                "Collage"
            }
        };
        private const string QuickFilterAll = "all";
        private const string QuickFilterInProgress = "in-progress";
        private const string QuickFilterReports = "reports";
        private const string QuickFilterDone = "done";
        private const string QuickFilterArchived = "archived";
        private const string EssaiFilterAll = "all";
        private const string EssaiFilterToProcess = "to-process";
        private const string EssaiFilterInProgress = "in-progress";
        private const string EssaiFilterDone = "done";
        private const string EssaisSortAlphabetical = "alphabetical";
        private const string EssaisSortStartDate = "start-date";
        private const string EssaisSortStatus = "status";
        private const string EssaisSortRemainingTests = "remaining-tests";
        private const double AgendaHoursPerDay = 7.0;
        private const double AgendaHourSlotHeight = 48.0;
        private const int AgendaSnapMinutes = 15;
        private const double AgendaScheduleEpsilon = 0.0001;
        private static readonly TimeSpan AgendaDisplayStartTime = TimeSpan.FromHours(7);
        private static readonly TimeSpan AgendaDisplayEndTime = TimeSpan.FromHours(18);

        private static readonly Dictionary<string, double> AgendaDureesEssais = new(StringComparer.OrdinalIgnoreCase)
        {
            ["Cyclage thermique"] = 4,
            ["Traction 100m"] = 0.25,
            ["Statique Bending"] = 0.25,
            ["Dimensionnel"] = 0.25,
            ["Crush"] = 2,
            ["Cut through"] = 2,
            ["Kink"] = 0.25,
            ["Repeated bending"] = 0.5,
            ["Torsion"] = 0.25,
            ["Abrasion marquage"] = 0.25,
            ["Abrasion gaine"] = 0.25,
            ["Friction gaine"] = 0.5,
            ["Traction pince"] = 2,
            ["Traction spiralé"] = 2,
            ["Pénétration d'eau"] = 7,
            ["Petite flamme"] = 0.25,
            ["Vibration éolienne"] = 14,
            ["Collage"] = 0.25
        };

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
            ["Traction verticale"] = new List<string> { "\u00C0 faire", "Fait", "Non concern\u00E9" },
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
            ["Exposition UV"] = new List<string> { "\u00C0 faire", "En cours", "Fait", "Non concern\u00E9" },
            ["CPR"] = new List<string> { "\u00C0 faire", "En cours", "Fait", "Non concern\u00E9" },
        };

        private string? searchNomProduit;
        private string? searchNomProduitEssais;
        private string? agendaSearchText;
        private SearchFieldOption? selectedProjetSearchField;
        private SearchFieldOption? selectedProjetEssaisSortOption;
        private bool isProjetEssaisSortDescending;
        private string activeQuickFilter = QuickFilterAll;
        private string activeEssaiFilter = EssaiFilterAll;
        private int totalProjets;
        private int statutsEnCours;
        private int rapportsEnCours;
        private int projetsFaits;
        private int projetsArchives;
        private Projet? selectedProjetEssais;
        private int essaisSelectionATraiter;
        private int essaisSelectionEnCours;
        private int essaisSelectionTermines;
        private int essaisSelectionTotal;
        private int essaisSelectionOk;
        private int essaisSelectionNok;
        private bool showDoneOkEssais = true;
        private string? selectedEssaisDefaultFamille;
        private IReadOnlyList<ProjetEnCoursSummary> projetsEnCoursDashboard = Array.Empty<ProjetEnCoursSummary>();
        private int projetsEnCoursDashboardCount;
        private int projetsEnAttentionDashboardCount;
        private int essaisTerminesDashboardTotal;
        private int essaisEnCoursDashboardTotal;
        private int essaisRestantsDashboardTotal;
        private int avancementGlobalEnCoursPourcentage;
        private string agendaWeekTitle = string.Empty;
        private readonly Dictionary<string, HashSet<string>> essaisDefaultsByFamille = new(StringComparer.OrdinalIgnoreCase);
        private readonly Stack<AgendaUndoSnapshot> agendaUndoSnapshots = new();
        private readonly Stack<EssaisUndoSnapshot> essaisUndoSnapshots = new();
        private readonly Stack<ProjectTableUndoSnapshot> projectTableUndoSnapshots = new();
        private bool isRestoringAgendaState;

        public event PropertyChangedEventHandler? PropertyChanged;

        public ObservableCollection<Projet> Projets { get; }
        public ICollectionView ProjetsView { get; }
        public ICollectionView ProjetsEssaisView { get; }

        public ObservableCollection<string> Clients { get; }
        public ObservableCollection<string> Demandeurs { get; }
        public ObservableCollection<string> FamillesProduit { get; }
        public ObservableCollection<string> TypesActivite { get; }
        public List<string> Statuts { get; }
        public List<string> NomsEssais { get; }
        public List<SearchFieldOption> ProjetSearchFields { get; }
        public List<SearchFieldOption> ProjetEssaisSortOptions { get; }
        public ObservableCollection<AgendaTaskItem> AgendaBacklogTasks { get; }
        public ObservableCollection<AgendaWorkDay> AgendaWeekDays { get; }
        public ObservableCollection<EssaiDefaultSelectionItem> EssaisParDefautSelection { get; } = new();

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

        public SearchFieldOption? SelectedProjetSearchField
        {
            get => selectedProjetSearchField;
            set
            {
                if (ReferenceEquals(selectedProjetSearchField, value))
                {
                    return;
                }

                selectedProjetSearchField = value;
                ProjetsView.Refresh();
                OnPropertyChanged(nameof(SelectedProjetSearchField));
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

        public SearchFieldOption? SelectedProjetEssaisSortOption
        {
            get => selectedProjetEssaisSortOption;
            set
            {
                if (ReferenceEquals(selectedProjetEssaisSortOption, value))
                {
                    return;
                }

                selectedProjetEssaisSortOption = value;
                ApplyProjetEssaisSorting();
                ProjetsEssaisView.Refresh();
                OnPropertyChanged(nameof(SelectedProjetEssaisSortOption));
                EnsureSelectedProjetEssais();
            }
        }

        public bool IsProjetEssaisSortDescending => isProjetEssaisSortDescending;
        public string ProjetEssaisSortDirectionGlyph => isProjetEssaisSortDescending ? "↓" : "↑";
        public string ProjetEssaisSortDirectionToolTip => isProjetEssaisSortDescending ? "Tri décroissant" : "Tri croissant";

        public bool IsAllFilterActive => activeQuickFilter == QuickFilterAll;
        public bool IsInProgressFilterActive => activeQuickFilter == QuickFilterInProgress;
        public bool IsReportsFilterActive => activeQuickFilter == QuickFilterReports;
        public bool IsDoneFilterActive => activeQuickFilter == QuickFilterDone;
        public bool IsArchivedFilterActive => activeQuickFilter == QuickFilterArchived;
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

        public int ProjetsArchives
        {
            get => projetsArchives;
            private set
            {
                if (projetsArchives == value)
                {
                    return;
                }

                projetsArchives = value;
                OnPropertyChanged(nameof(ProjetsArchives));
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

        public int EssaisSelectionOk
        {
            get => essaisSelectionOk;
            private set
            {
                if (essaisSelectionOk == value)
                {
                    return;
                }

                essaisSelectionOk = value;
                OnPropertyChanged(nameof(EssaisSelectionOk));
            }
        }

        public int EssaisSelectionNok
        {
            get => essaisSelectionNok;
            private set
            {
                if (essaisSelectionNok == value)
                {
                    return;
                }

                essaisSelectionNok = value;
                OnPropertyChanged(nameof(EssaisSelectionNok));
            }
        }

        public bool ShowDoneOkEssais => showDoneOkEssais;

        public string ToggleDoneOkEssaisLabel => showDoneOkEssais ? "Masquer les essais terminés" : "Afficher les essais terminés";

        public string? SelectedEssaisDefaultFamille
        {
            get => selectedEssaisDefaultFamille;
            set
            {
                if (selectedEssaisDefaultFamille == value)
                {
                    return;
                }

                selectedEssaisDefaultFamille = value;
                RefreshEssaisParDefautSelection();
                OnPropertyChanged(nameof(SelectedEssaisDefaultFamille));
                OnPropertyChanged(nameof(EssaisParDefautSummary));
            }
        }

        public string EssaisParDefautSummary =>
            $"{EssaisParDefautSelection.Count(item => item.IsSelected)} essai(s) cochés par défaut";

        public IReadOnlyList<ProjetEnCoursSummary> ProjetsEnCoursDashboard => projetsEnCoursDashboard;

        public bool HasProjetsEnCoursDashboard => projetsEnCoursDashboardCount > 0;

        public int ProjetsEnCoursDashboardCount => projetsEnCoursDashboardCount;

        public int ProjetsEnAttentionDashboardCount => projetsEnAttentionDashboardCount;

        public int EssaisTerminesDashboardTotal => essaisTerminesDashboardTotal;

        public int EssaisEnCoursDashboardTotal => essaisEnCoursDashboardTotal;

        public int EssaisRestantsDashboardTotal => essaisRestantsDashboardTotal;

        public int AvancementGlobalEnCoursPourcentage => avancementGlobalEnCoursPourcentage;

        public string AvancementGlobalEnCoursLabel =>
            (essaisTerminesDashboardTotal + essaisEnCoursDashboardTotal + essaisRestantsDashboardTotal) == 0
                ? "Aucun essai concerné"
                : $"{essaisTerminesDashboardTotal} essais terminés sur {essaisTerminesDashboardTotal + essaisEnCoursDashboardTotal + essaisRestantsDashboardTotal}";

        public IEnumerable<EssaiSuivi> EssaisPreQualificationFiltres =>
            GetFilteredEssais(SelectedProjetEssais?.EssaisPreQualification ?? Enumerable.Empty<EssaiSuivi>());

        public IEnumerable<EssaiSuivi> EssaisQualificationFiltres =>
            GetFilteredEssais(SelectedProjetEssais?.EssaisQualification ?? Enumerable.Empty<EssaiSuivi>());

        public IEnumerable<EssaiSuivi> EssaisExterieursFiltres =>
            GetFilteredEssais(SelectedProjetEssais?.EssaisExterieurs ?? Enumerable.Empty<EssaiSuivi>());

        public bool HasFilteredEssais =>
            EssaisPreQualificationFiltres.Any() ||
            EssaisQualificationFiltres.Any() ||
            EssaisExterieursFiltres.Any();

        public string EmptyEssaisFilterMessage => activeEssaiFilter switch
        {
            EssaiFilterToProcess => "Aucun essai à faire.",
            EssaiFilterInProgress => "Aucun essai en cours.",
            EssaiFilterDone => "Aucun essai terminé.",
            _ => "Aucun résultat."
        };

        public string AgendaWeekTitle
        {
            get => agendaWeekTitle;
            private set
            {
                if (agendaWeekTitle == value)
                {
                    return;
                }

                agendaWeekTitle = value;
                OnPropertyChanged(nameof(AgendaWeekTitle));
            }
        }

        public string? AgendaSearchText
        {
            get => agendaSearchText;
            set
            {
                if (agendaSearchText == value)
                {
                    return;
                }

                agendaSearchText = value;
                OnPropertyChanged(nameof(AgendaSearchText));
                RefreshAgendaCollections();
            }
        }

        public IEnumerable<AgendaTaskItem> AgendaBacklogVisibleTasks =>
            GetUnplannedAgendaTasks()
                .Where(MatchesAgendaSearch)
                .OrderBy(task => NormalizeText(task.NumeroProjet), StringComparer.Ordinal)
                .ThenBy(task => NormalizeText(task.NomProduit), StringComparer.Ordinal)
                .ThenBy(task => NormalizeText(task.NomEssai), StringComparer.Ordinal)
                .ToList();

        public bool CanUndoAgenda => agendaUndoSnapshots.Count > 0;
        public bool CanUndoEssaisBulkAction => essaisUndoSnapshots.Count > 0;
        public bool CanUndoProjectTable => projectTableUndoSnapshots.Count > 0;

        public int AgendaBacklogVisibleCount => AgendaBacklogVisibleTasks.Count();

        public int AgendaBacklogRemainingCount => GetUnplannedAgendaTasks().Count();

        public bool HasAgendaSearch => !string.IsNullOrWhiteSpace(AgendaSearchText);

        public bool HasHiddenAgendaBacklogTasks => HasAgendaSearch && AgendaBacklogVisibleCount < AgendaBacklogRemainingCount;

        public string AgendaBacklogFilterSummary => HasHiddenAgendaBacklogTasks
            ? $"{AgendaBacklogVisibleCount}/{AgendaBacklogRemainingCount} affiché(s)"
            : $"{AgendaBacklogRemainingCount} restant(s)";

        public int AgendaPlannedTaskCount => AgendaWeekDays.Sum(day => day.PlannedTasks.Count);

        public int AgendaOverflowTaskCount => AgendaWeekDays
            .SelectMany(day => day.PlannedTasks)
            .Distinct()
            .Count(task => task.IsOverflow);

        public string AgendaPlannedDurationLabel
        {
            get
            {
                double totalHours = AgendaWeekDays
                    .SelectMany(day => day.PlannedTasks)
                    .Distinct()
                    .Sum(task => task.DureeHeures);

                return AgendaDurationFormatter.Format(totalHours, totalHours / AgendaHoursPerDay);
            }
        }

        public IEnumerable<AgendaHourMarker> AgendaDisplayHourMarkers
        {
            get
            {
                TimeSpan current = AgendaDisplayStartTime;
                double offset = 0;

                while (current <= AgendaDisplayEndTime)
                {
                    yield return new AgendaHourMarker($"{current:hh\\:mm}", offset);
                    current = current.Add(TimeSpan.FromHours(1));
                    offset += AgendaHourSlotHeight;
                }
            }
        }

        public IEnumerable<string> AgendaDisplayGridSlots
        {
            get
            {
                TimeSpan current = AgendaDisplayStartTime;
                while (current < AgendaDisplayEndTime)
                {
                    yield return $"{current:hh\\:mm}";
                    current = current.Add(TimeSpan.FromHours(1));
                }
            }
        }

        public double AgendaTimelineHeight => (AgendaDisplayEndTime - AgendaDisplayStartTime).TotalHours * AgendaHourSlotHeight;

        public void AjouterProjet(Projet projet)
        {
            EnsureEssaisForProject(projet);
            Projets.Add(projet);
            SelectedProjetEssais = projet;
        }

        public bool ProjetExiste(string? numeroProjet, string? nomProduit)
        {
            string numeroNormalise = NormalizeText(numeroProjet);
            string produitNormalise = NormalizeText(nomProduit);

            return Projets.Any(projet =>
                NormalizeText(projet.NumeroProjet) == numeroNormalise &&
                NormalizeText(projet.NomProduit) == produitNormalise);
        }

        public bool ProjetExisteAutre(Projet projetCourant, string? numeroProjet, string? nomProduit)
        {
            string numeroNormalise = NormalizeText(numeroProjet);
            string produitNormalise = NormalizeText(nomProduit);

            return Projets.Any(projet =>
                !ReferenceEquals(projet, projetCourant) &&
                NormalizeText(projet.NumeroProjet) == numeroNormalise &&
                NormalizeText(projet.NomProduit) == produitNormalise);
        }

        public void SupprimerProjet(Projet projet)
        {
            if (!Projets.Contains(projet))
            {
                return;
            }

            if (ReferenceEquals(SelectedProjetEssais, projet))
            {
                SelectedProjetEssais = null;
            }

            Projets.Remove(projet);
        }

        public bool ToggleProjetArchive(Projet projet)
        {
            if (!Projets.Contains(projet))
            {
                return false;
            }

            if (!projet.EstArchive && !projet.PeutArchiver)
            {
                return false;
            }

            SaveProjectTableUndoSnapshot();
            projet.EstArchive = !projet.EstArchive;
            return true;
        }

        public void SaveProjectTableUndoSnapshot()
        {
            ProjectTableUndoSnapshot snapshot = CaptureProjectTableSnapshot();

            if (projectTableUndoSnapshots.Count > 0 &&
                projectTableUndoSnapshots.Peek().States.SequenceEqual(snapshot.States))
            {
                return;
            }

            projectTableUndoSnapshots.Push(snapshot);

            while (projectTableUndoSnapshots.Count > 30)
            {
                ProjectTableUndoSnapshot[] preserved = projectTableUndoSnapshots.Reverse().Take(30).ToArray();
                projectTableUndoSnapshots.Clear();

                for (int i = preserved.Length - 1; i >= 0; i--)
                {
                    projectTableUndoSnapshots.Push(preserved[i]);
                }
            }

            OnPropertyChanged(nameof(CanUndoProjectTable));
        }

        public void UndoProjectTableLastAction()
        {
            if (projectTableUndoSnapshots.Count == 0)
            {
                return;
            }

            ProjectTableUndoSnapshot snapshot = projectTableUndoSnapshots.Pop();
            RestoreProjectTableSnapshot(snapshot);
            OnPropertyChanged(nameof(CanUndoProjectTable));
        }

        public void MarkAllEssaisDoneAndOk(Projet projet)
        {
            if (!Projets.Contains(projet))
            {
                return;
            }

            if (!SaveEssaisUndoSnapshotIfNeeded(
                    projet,
                    essai => essai.EstConcerne &&
                             (!string.Equals(essai.Statut, "Fait", StringComparison.Ordinal) ||
                              !string.Equals(essai.ResultatTraitement, "OK", StringComparison.Ordinal))))
            {
                return;
            }

            foreach (EssaiSuivi essai in projet.Essais.Where(essai => essai.EstConcerne))
            {
                essai.Statut = "Fait";
                essai.ResultatTraitement = "OK";
            }

            RefreshEssaiCollections();
            RefreshSelectedProjectStatistics();
            RefreshAgendaTasks();
        }

        public void MarkAllEssaisToDo(Projet projet)
        {
            if (!Projets.Contains(projet))
            {
                return;
            }

            if (!SaveEssaisUndoSnapshotIfNeeded(
                    projet,
                    essai => essai.EstConcerne &&
                             (!string.Equals(essai.Statut, "\u00C0 faire", StringComparison.Ordinal) ||
                              essai.ResultatTraitement != null)))
            {
                return;
            }

            foreach (EssaiSuivi essai in projet.Essais.Where(essai => essai.EstConcerne))
            {
                essai.Statut = "\u00C0 faire";
                essai.ResultatTraitement = null;
            }

            RefreshEssaiCollections();
            RefreshSelectedProjectStatistics();
            RefreshAgendaTasks();
        }

        public void UndoLastEssaisBulkAction()
        {
            if (essaisUndoSnapshots.Count == 0)
            {
                return;
            }

            EssaisUndoSnapshot snapshot = essaisUndoSnapshots.Pop();
            OnPropertyChanged(nameof(CanUndoEssaisBulkAction));

            RestoreEssaisSnapshot(snapshot);
        }

        public int CountProjectsWithClient(string? client)
        {
            return CountProjectsByReference(client, projet => projet.Client);
        }

        public int CountProjectsWithDemandeur(string? demandeur)
        {
            return CountProjectsByReference(demandeur, projet => projet.Demandeur);
        }

        public bool AddClient(string? value, out string normalizedValue, out string message)
        {
            return AddReferenceValue(Clients, nameof(Clients), "client", value, out normalizedValue, out message);
        }

        public bool RenameClient(string? currentValue, string? newValue, out string normalizedValue, out string message)
        {
            return RenameReferenceValue(
                Clients,
                nameof(Clients),
                "client",
                currentValue,
                newValue,
                projet => projet.Client,
                (projet, value) => projet.Client = value,
                out normalizedValue,
                out message);
        }

        public bool DeleteClient(string? value, out string message)
        {
            string? fallbackValue = Clients.FirstOrDefault(item =>
                NormalizeText(item) == "n/a" &&
                NormalizeText(item) != NormalizeText(value));

            fallbackValue ??= Clients.FirstOrDefault(item =>
                NormalizeText(item) != NormalizeText(value));

            return DeleteReferenceValue(
                Clients,
                nameof(Clients),
                "client",
                value,
                fallbackValue,
                projet => projet.Client,
                (projet, replacement) => projet.Client = replacement,
                out message);
        }

        public bool AddDemandeur(string? value, out string normalizedValue, out string message)
        {
            return AddReferenceValue(Demandeurs, nameof(Demandeurs), "demandeur", value, out normalizedValue, out message);
        }

        public bool RenameDemandeur(string? currentValue, string? newValue, out string normalizedValue, out string message)
        {
            return RenameReferenceValue(
                Demandeurs,
                nameof(Demandeurs),
                "demandeur",
                currentValue,
                newValue,
                projet => projet.Demandeur,
                (projet, value) => projet.Demandeur = value,
                out normalizedValue,
                out message);
        }

        public bool DeleteDemandeur(string? value, out string message)
        {
            return DeleteReferenceValue(
                Demandeurs,
                nameof(Demandeurs),
                "demandeur",
                value,
                null,
                projet => projet.Demandeur,
                (projet, replacement) => projet.Demandeur = replacement,
                out message);
        }

        public IReadOnlyCollection<string> GetDefaultEssaisForFamille(string? famille)
        {
            if (!string.IsNullOrWhiteSpace(famille) &&
                essaisDefaultsByFamille.TryGetValue(famille, out HashSet<string>? essais))
            {
                return essais.ToList();
            }

            return NomsEssais.ToList();
        }

        public void SetAllEssaisParDefautSelection(bool isSelected)
        {
            foreach (EssaiDefaultSelectionItem item in EssaisParDefautSelection)
            {
                item.IsSelected = isSelected;
            }

            OnPropertyChanged(nameof(EssaisParDefautSummary));
        }

        public bool SaveSelectedEssaisParDefaut(out string message)
        {
            if (string.IsNullOrWhiteSpace(SelectedEssaisDefaultFamille))
            {
                message = "Sélectionne une famille de produits.";
                return false;
            }

            essaisDefaultsByFamille[SelectedEssaisDefaultFamille] = EssaisParDefautSelection
                .Where(item => item.IsSelected)
                .Select(item => item.NomEssai)
                .ToHashSet(StringComparer.OrdinalIgnoreCase);

            SaveReferenceProperties();
            OnPropertyChanged(nameof(EssaisParDefautSummary));

            message = $"Essais par défaut enregistrés pour {SelectedEssaisDefaultFamille}.";
            return true;
        }

        public MainViewModel()
        {
            ReferencePropertiesSettings? referenceSettings = LoadReferenceProperties();

            Clients = BuildReferenceValues(DefaultClients, referenceSettings?.Clients);
            Demandeurs = BuildReferenceValues(DefaultDemandeurs, referenceSettings?.Demandeurs);
            FamillesProduit = BuildReferenceValues(DefaultFamillesProduit, referenceSettings?.FamillesProduit);
            TypesActivite = BuildReferenceValues(DefaultTypesActivite, referenceSettings?.TypesActivite);
            Statuts = new List<string>
            {
                "À faire",
                "Pré-qualification en cours",
                "Qualification en cours",
                "Rapport en cours",
                "Rapport terminé",
                "Fait"
            };
            NomsEssais = new List<string>(DefaultNomsEssais);
            InitializeEssaisDefaults(referenceSettings?.EssaisParFamille);
            selectedEssaisDefaultFamille = FamillesProduit.FirstOrDefault();
            RefreshEssaisParDefautSelection();
            ProjetSearchFields = new List<SearchFieldOption>
            {
                new("NomProduit", "Produit"),
                new("NumeroProjet", "Projet")
            };
            selectedProjetSearchField = ProjetSearchFields[0];
            ProjetEssaisSortOptions = new List<SearchFieldOption>
            {
                new(EssaisSortAlphabetical, "A-Z"),
                new(EssaisSortStartDate, "Date début"),
                new(EssaisSortStatus, "Statut"),
                new(EssaisSortRemainingTests, "Restants")
            };
            selectedProjetEssaisSortOption = ProjetEssaisSortOptions[0];
            AgendaBacklogTasks = new ObservableCollection<AgendaTaskItem>();
            AgendaWeekDays = CreateAgendaWeekDays();

            Projets = ChargerProjets();

            foreach (Projet projet in Projets)
            {
                EnsureEssaisForProject(projet);
            }

            ProjetsView = new CollectionViewSource { Source = Projets }.View;
            ProjetsView.Filter = FilterProjet;

            ProjetsEssaisView = new CollectionViewSource { Source = Projets }.View;
            ProjetsEssaisView.Filter = FilterProjetEssais;
            ApplyProjetEssaisSorting();

            Projets.CollectionChanged += Projets_CollectionChanged;

            foreach (Projet projet in Projets)
            {
                AttacherProjet(projet);
            }

            RefreshStatistics();
            EnsureSelectedProjetEssais();
            RefreshEnCoursDashboard();
            RefreshAgendaTasks();
        }

        private static ReferencePropertiesSettings? LoadReferenceProperties()
        {
            if (!File.Exists(ReferencePropertiesPath))
            {
                return null;
            }

            try
            {
                string json = File.ReadAllText(ReferencePropertiesPath, Encoding.UTF8);
                return JsonSerializer.Deserialize<ReferencePropertiesSettings>(json);
            }
            catch
            {
                return null;
            }
        }

        private static ObservableCollection<string> BuildReferenceValues(
            IEnumerable<string> fallbackValues,
            IEnumerable<string>? savedValues)
        {
            IEnumerable<string> sourceValues = savedValues ?? fallbackValues;
            var values = new ObservableCollection<string>();

            foreach (string value in sourceValues)
            {
                string normalizedValue = NormalizeReferenceValue(value);
                if (string.IsNullOrWhiteSpace(normalizedValue) ||
                    values.Any(existing => string.Equals(NormalizeText(existing), NormalizeText(normalizedValue), StringComparison.Ordinal)))
                {
                    continue;
                }

                values.Add(normalizedValue);
            }

            return values;
        }

        private void InitializeEssaisDefaults(Dictionary<string, List<string>>? savedValues)
        {
            essaisDefaultsByFamille.Clear();

            foreach (string famille in FamillesProduit)
            {
                IEnumerable<string> essais = GetSavedEssaisDefaultsForFamille(savedValues, famille) ??
                                             GetBuiltInEssaisDefaultsForFamille(famille);

                essaisDefaultsByFamille[famille] = essais
                    .Where(nomEssai => NomsEssais.Any(known => string.Equals(known, nomEssai, StringComparison.OrdinalIgnoreCase)))
                    .ToHashSet(StringComparer.OrdinalIgnoreCase);
            }
        }

        private static IEnumerable<string>? GetSavedEssaisDefaultsForFamille(
            Dictionary<string, List<string>>? savedValues,
            string famille)
        {
            if (savedValues == null)
            {
                return null;
            }

            KeyValuePair<string, List<string>> match = savedValues.FirstOrDefault(item =>
                NormalizeText(item.Key) == NormalizeText(famille));

            return string.IsNullOrWhiteSpace(match.Key) ? null : match.Value;
        }

        private static IEnumerable<string> GetBuiltInEssaisDefaultsForFamille(string famille)
        {
            return DefaultEssaisParFamille.TryGetValue(famille, out string[]? essais)
                ? essais
                : DefaultNomsEssais;
        }

        private void RefreshEssaisParDefautSelection()
        {
            foreach (EssaiDefaultSelectionItem item in EssaisParDefautSelection)
            {
                item.PropertyChanged -= EssaiDefaultSelectionItem_PropertyChanged;
            }

            EssaisParDefautSelection.Clear();
            IReadOnlyCollection<string> selectedEssais = GetDefaultEssaisForFamille(SelectedEssaisDefaultFamille);

            foreach (string nomEssai in NomsEssais)
            {
                var item = new EssaiDefaultSelectionItem(
                    nomEssai,
                    selectedEssais.Any(selected => string.Equals(selected, nomEssai, StringComparison.OrdinalIgnoreCase)));
                item.PropertyChanged += EssaiDefaultSelectionItem_PropertyChanged;
                EssaisParDefautSelection.Add(item);
            }

            OnPropertyChanged(nameof(EssaisParDefautSummary));
        }

        private void EssaiDefaultSelectionItem_PropertyChanged(object? sender, PropertyChangedEventArgs e)
        {
            if (e.PropertyName == nameof(EssaiDefaultSelectionItem.IsSelected))
            {
                OnPropertyChanged(nameof(EssaisParDefautSummary));
            }
        }

        private bool AddReferenceValue(
            ObservableCollection<string> values,
            string propertyName,
            string label,
            string? value,
            out string normalizedValue,
            out string message)
        {
            normalizedValue = NormalizeReferenceValue(value);

            if (string.IsNullOrWhiteSpace(normalizedValue))
            {
                message = $"Le {label} ne peut pas être vide.";
                return false;
            }

            if (FindReferenceIndex(values, normalizedValue) >= 0)
            {
                message = $"Ce {label} existe déjà.";
                return false;
            }

            values.Add(normalizedValue);
            SaveReferenceProperties();
            OnPropertyChanged(propertyName);
            message = $"{Capitalize(label)} ajouté.";
            return true;
        }

        private bool RenameReferenceValue(
            ObservableCollection<string> values,
            string propertyName,
            string label,
            string? currentValue,
            string? newValue,
            Func<Projet, string?> projectSelector,
            Action<Projet, string?> projectUpdater,
            out string normalizedValue,
            out string message)
        {
            normalizedValue = NormalizeReferenceValue(newValue);
            int currentIndex = FindReferenceIndex(values, currentValue);

            if (currentIndex < 0)
            {
                message = $"Sélectionne un {label} à modifier.";
                return false;
            }

            if (string.IsNullOrWhiteSpace(normalizedValue))
            {
                message = $"Le {label} ne peut pas être vide.";
                return false;
            }

            int duplicateIndex = FindReferenceIndex(values, normalizedValue);
            if (duplicateIndex >= 0 && duplicateIndex != currentIndex)
            {
                message = $"Ce {label} existe déjà.";
                return false;
            }

            string oldValue = values[currentIndex];
            values[currentIndex] = normalizedValue;
            int updatedProjects = UpdateProjectsReference(oldValue, normalizedValue, projectSelector, projectUpdater);

            SaveReferenceProperties();
            RefreshAfterReferenceEdit(propertyName, updatedProjects > 0);

            message = updatedProjects == 0
                ? $"{Capitalize(label)} modifié."
                : $"{Capitalize(label)} modifié dans {updatedProjects} projet(s).";
            return true;
        }

        private bool DeleteReferenceValue(
            ObservableCollection<string> values,
            string propertyName,
            string label,
            string? value,
            string? replacementValue,
            Func<Projet, string?> projectSelector,
            Action<Projet, string?> projectUpdater,
            out string message)
        {
            int index = FindReferenceIndex(values, value);

            if (index < 0)
            {
                message = $"Sélectionne un {label} à supprimer.";
                return false;
            }

            string removedValue = values[index];
            values.RemoveAt(index);
            int updatedProjects = UpdateProjectsReference(removedValue, replacementValue, projectSelector, projectUpdater);

            SaveReferenceProperties();
            RefreshAfterReferenceEdit(propertyName, updatedProjects > 0);

            message = updatedProjects == 0
                ? $"{Capitalize(label)} supprimé."
                : $"{Capitalize(label)} supprimé et retiré de {updatedProjects} projet(s).";
            return true;
        }

        private int UpdateProjectsReference(
            string oldValue,
            string? newValue,
            Func<Projet, string?> selector,
            Action<Projet, string?> updater)
        {
            int updatedProjects = 0;
            string normalizedOldValue = NormalizeText(oldValue);

            foreach (Projet projet in Projets.Where(projet => NormalizeText(selector(projet)) == normalizedOldValue))
            {
                updater(projet, newValue);
                updatedProjects++;
            }

            return updatedProjects;
        }

        private int CountProjectsByReference(string? value, Func<Projet, string?> selector)
        {
            string normalizedValue = NormalizeText(value);
            if (string.IsNullOrWhiteSpace(normalizedValue))
            {
                return 0;
            }

            return Projets.Count(projet => NormalizeText(selector(projet)) == normalizedValue);
        }

        private void RefreshAfterReferenceEdit(string propertyName, bool projectsWereUpdated)
        {
            OnPropertyChanged(propertyName);

            if (!projectsWereUpdated)
            {
                return;
            }

            Sauvegarder();
            ProjetsView.Refresh();
            ProjetsEssaisView.Refresh();
            EnsureSelectedProjetEssais();
            RefreshStatistics();
            RefreshSelectedProjectStatistics();
            RefreshAgendaTasks();
        }

        private void SaveReferenceProperties()
        {
            var settings = new ReferencePropertiesSettings
            {
                Clients = Clients.ToList(),
                Demandeurs = Demandeurs.ToList(),
                FamillesProduit = FamillesProduit.ToList(),
                TypesActivite = TypesActivite.ToList(),
                EssaisParFamille = essaisDefaultsByFamille.ToDictionary(
                    item => item.Key,
                    item => NomsEssais
                        .Where(nomEssai => item.Value.Contains(nomEssai))
                        .ToList(),
                    StringComparer.OrdinalIgnoreCase)
            };

            string json = JsonSerializer.Serialize(
                settings,
                new JsonSerializerOptions
                {
                    WriteIndented = true,
                    Encoder = JavaScriptEncoder.UnsafeRelaxedJsonEscaping
                });

            File.WriteAllText(ReferencePropertiesPath, json, new UTF8Encoding(false));
        }

        private static int FindReferenceIndex(ObservableCollection<string> values, string? value)
        {
            string normalizedValue = NormalizeText(value);

            for (int index = 0; index < values.Count; index++)
            {
                if (NormalizeText(values[index]) == normalizedValue)
                {
                    return index;
                }
            }

            return -1;
        }

        private static string NormalizeReferenceValue(string? value)
        {
            return (value ?? string.Empty).Trim();
        }

        private static string Capitalize(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                return value;
            }

            return char.ToUpper(value[0], CultureInfo.GetCultureInfo("fr-FR")) + value[1..];
        }

        private ObservableCollection<Projet> ChargerProjets()
        {
            if (!File.Exists(DataFilePath))
            {
                return new ObservableCollection<Projet>();
            }

            string json = File.ReadAllText(DataFilePath, Encoding.UTF8);
            ObservableCollection<Projet> projets = JsonSerializer.Deserialize<ObservableCollection<Projet>>(json) ?? new ObservableCollection<Projet>();
            bool hasRepairs = false;

            foreach (Projet projet in projets)
            {
                hasRepairs |= RepairLoadedProject(projet);
            }

            if (hasRepairs)
            {
                string repairedJson = JsonSerializer.Serialize(
                    projets,
                    new JsonSerializerOptions
                    {
                        WriteIndented = true,
                        Encoder = JavaScriptEncoder.UnsafeRelaxedJsonEscaping
                    });

                File.WriteAllText(DataFilePath, repairedJson, new UTF8Encoding(false));
            }

            return projets;
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

            if (projet.EstArchive)
            {
                return false;
            }

            if (string.IsNullOrWhiteSpace(SearchNomProduitEssais))
            {
                return true;
            }

            string searchValue = NormalizeText(SearchNomProduitEssais);

            return NormalizeText(projet.NomProduit).Contains(searchValue) ||
                   NormalizeText(projet.NumeroProjet).Contains(searchValue);
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
            RefreshAgendaTasks();
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
            if (e.PropertyName == nameof(Projet.NomProduit) ||
                e.PropertyName == nameof(Projet.NumeroProjet) ||
                e.PropertyName == nameof(Projet.Statut) ||
                e.PropertyName == nameof(Projet.EstArchive) ||
                e.PropertyName == nameof(Projet.DateDebut) ||
                e.PropertyName == nameof(Projet.DateDebutValue))
            {
                ProjetsView.Refresh();
                ProjetsEssaisView.Refresh();
                EnsureSelectedProjetEssais();
            }

            Sauvegarder();
            RefreshStatistics();
            RefreshSelectedProjectStatistics();
            RefreshAgendaTasks();
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

            if (sender is ObservableCollection<EssaiSuivi> essaisCollection)
            {
                Projet? projet = Projets.FirstOrDefault(item => ReferenceEquals(item.Essais, essaisCollection));
                if (projet != null)
                {
                    UpdateProjectStatusFromEssais(projet);
                }
            }

            Sauvegarder();
            ProjetsEssaisView.Refresh();
            EnsureSelectedProjetEssais();
            RefreshSelectedProjectStatistics();
            RefreshAgendaTasks();
        }

        private void Essai_PropertyChanged(object? sender, PropertyChangedEventArgs e)
        {
            if (sender is EssaiSuivi essai)
            {
                Projet? projet = Projets.FirstOrDefault(item => item.Essais.Contains(essai));
                if (projet != null)
                {
                    UpdateProjectStatusFromEssais(projet);
                }
            }

            Sauvegarder();
            ProjetsEssaisView.Refresh();
            EnsureSelectedProjetEssais();
            RefreshSelectedProjectStatistics();
            RefreshAgendaTasks();
        }

        private IEnumerable<Projet> GetSearchScopedProjects()
        {
            return Projets.Where(MatchesSearch).Where(projet => !projet.EstArchive);
        }

        private IEnumerable<Projet> GetSearchScopedArchivedProjects()
        {
            return Projets.Where(MatchesSearch).Where(projet => projet.EstArchive);
        }

        private void RefreshStatistics()
        {
            TotalProjets = GetSearchScopedProjects().Count();
            StatutsEnCours = CountStatusContaining("cours");
            RapportsEnCours = CountStatusContaining("rapport");
            ProjetsFaits = GetSearchScopedProjects().Count(projet => projet.EstTermine);
            ProjetsArchives = GetSearchScopedArchivedProjects().Count();
            OnPropertyChanged(nameof(IsAllFilterActive));
            OnPropertyChanged(nameof(IsInProgressFilterActive));
            OnPropertyChanged(nameof(IsReportsFilterActive));
            OnPropertyChanged(nameof(IsDoneFilterActive));
            OnPropertyChanged(nameof(IsArchivedFilterActive));
            RefreshEnCoursDashboard();
        }

        private void RefreshSelectedProjectStatistics()
        {
            if (SelectedProjetEssais == null)
            {
                EssaisSelectionTotal = 0;
                EssaisSelectionEnCours = 0;
                EssaisSelectionATraiter = 0;
                EssaisSelectionTermines = 0;
                EssaisSelectionOk = 0;
                EssaisSelectionNok = 0;
                RefreshEssaiCollections();
                RefreshEnCoursDashboard();
                return;
            }

            IEnumerable<EssaiSuivi> essaisConcernes = SelectedProjetEssais.Essais.Where(essai => essai.EstConcerne);
            EssaisSelectionTotal = essaisConcernes.Count();
            EssaisSelectionATraiter = essaisConcernes.Count(IsEssaiToProcess);
            EssaisSelectionEnCours = essaisConcernes.Count(IsEssaiInProgress);
            EssaisSelectionTermines = essaisConcernes.Count(IsEssaiDone);
            EssaisSelectionOk = essaisConcernes.Count(essai => IsEssaiDone(essai) && NormalizeText(essai.ResultatTraitement) == "ok");
            EssaisSelectionNok = essaisConcernes.Count(essai => IsEssaiDone(essai) && NormalizeText(essai.ResultatTraitement) == "nok");
            RefreshEssaiCollections();
            RefreshEnCoursDashboard();
        }

        public void ToggleProjetEssaisSortDirection()
        {
            isProjetEssaisSortDescending = !isProjetEssaisSortDescending;
            ApplyProjetEssaisSorting();
            ProjetsEssaisView.Refresh();
            OnPropertyChanged(nameof(IsProjetEssaisSortDescending));
            OnPropertyChanged(nameof(ProjetEssaisSortDirectionGlyph));
            OnPropertyChanged(nameof(ProjetEssaisSortDirectionToolTip));
            EnsureSelectedProjetEssais();
        }

        private int CountStatusContaining(string expectedPart)
        {
            string expected = NormalizeText(expectedPart);
            return GetSearchScopedProjects().Count(projet => NormalizeText(projet.Statut).Contains(expected));
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

        public void SetQuickFilterToArchived()
        {
            SetQuickFilter(QuickFilterArchived);
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

        public void ToggleDoneOkEssaisVisibility()
        {
            showDoneOkEssais = !showDoneOkEssais;
            OnPropertyChanged(nameof(ShowDoneOkEssais));
            OnPropertyChanged(nameof(ToggleDoneOkEssaisLabel));
            RefreshEssaiCollections();
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

        private void ApplyProjetEssaisSorting()
        {
            if (ProjetsEssaisView is not ListCollectionView listCollectionView)
            {
                return;
            }

            listCollectionView.CustomSort = new DelegateComparer((left, right) =>
            {
                if (left is not Projet leftProjet || right is not Projet rightProjet)
                {
                    return 0;
                }

                return CompareProjetsEssais(leftProjet, rightProjet);
            });
        }

        private bool MatchesSearch(Projet projet)
        {
            if (string.IsNullOrWhiteSpace(SearchNomProduit))
            {
                return true;
            }

            string searchValue = NormalizeText(SearchNomProduit);

            return SelectedProjetSearchField?.Key switch
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
            if (activeQuickFilter == QuickFilterArchived)
            {
                return projet.EstArchive;
            }

            if (projet.EstArchive)
            {
                return false;
            }

            string statut = NormalizeText(projet.Statut);

            return activeQuickFilter switch
            {
                QuickFilterInProgress => statut.Contains("cours"),
                QuickFilterReports => statut.Contains("rapport"),
                QuickFilterDone => projet.EstTermine,
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
            if (!showDoneOkEssais && IsEssaiDoneAndOk(essai))
            {
                return false;
            }

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
            return statut.Contains("faire");
        }

        private static bool IsEssaiDone(EssaiSuivi essai)
        {
            return essai.EstConcerne && essai.ProgressionPourcentage == 100;
        }

        private static bool IsEssaiDoneAndOk(EssaiSuivi essai)
        {
            return IsEssaiDone(essai) && NormalizeText(essai.ResultatTraitement) == "ok";
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

        private int CompareProjetsEssais(Projet left, Projet right)
        {
            int comparison = (SelectedProjetEssaisSortOption?.Key ?? EssaisSortAlphabetical) switch
            {
                EssaisSortStartDate => CompareProjectStartDates(left, right),
                EssaisSortStatus => CompareProjectStatuses(left, right),
                EssaisSortRemainingTests => CompareRemainingEssais(left, right),
                _ => CompareProjectNames(left, right)
            };

            if (isProjetEssaisSortDescending)
            {
                comparison *= -1;
            }

            if (comparison != 0)
            {
                return comparison;
            }

            comparison = CompareProjectNames(left, right);
            if (comparison != 0)
            {
                return comparison;
            }

            return string.Compare(
                NormalizeText(left.NumeroProjet),
                NormalizeText(right.NumeroProjet),
                StringComparison.Ordinal);
        }

        private int CompareProjectNames(Projet left, Projet right)
        {
            return string.Compare(
                NormalizeText(left.NomProduit),
                NormalizeText(right.NomProduit),
                StringComparison.Ordinal);
        }

        private int CompareProjectStartDates(Projet left, Projet right)
        {
            DateTime? leftDate = left.DateDebutValue;
            DateTime? rightDate = right.DateDebutValue;

            if (leftDate.HasValue && rightDate.HasValue)
            {
                return DateTime.Compare(leftDate.Value, rightDate.Value);
            }

            if (leftDate.HasValue)
            {
                return -1;
            }

            if (rightDate.HasValue)
            {
                return 1;
            }

            return 0;
        }

        private int CompareProjectStatuses(Projet left, Projet right)
        {
            int leftIndex = GetProjectStatusSortIndex(left.Statut);
            int rightIndex = GetProjectStatusSortIndex(right.Statut);
            return leftIndex.CompareTo(rightIndex);
        }

        private int GetProjectStatusSortIndex(string? statut)
        {
            string normalizedStatus = NormalizeText(statut);

            for (int index = 0; index < Statuts.Count; index++)
            {
                if (NormalizeText(Statuts[index]) == normalizedStatus)
                {
                    return index;
                }
            }

            return int.MaxValue;
        }

        private int CompareRemainingEssais(Projet left, Projet right)
        {
            int leftRemaining = CountRemainingEssais(left);
            int rightRemaining = CountRemainingEssais(right);
            return rightRemaining.CompareTo(leftRemaining);
        }

        private static int CountRemainingEssais(Projet projet)
        {
            return projet.Essais.Count(essai => essai.EstConcerne && !IsEssaiDone(essai));
        }

        private void RefreshEssaiCollections()
        {
            OnPropertyChanged(nameof(EssaisPreQualificationFiltres));
            OnPropertyChanged(nameof(EssaisQualificationFiltres));
            OnPropertyChanged(nameof(EssaisExterieursFiltres));
            OnPropertyChanged(nameof(HasFilteredEssais));
            OnPropertyChanged(nameof(EmptyEssaisFilterMessage));
            OnPropertyChanged(nameof(IsEssaiToProcessFilterActive));
            OnPropertyChanged(nameof(IsEssaiInProgressFilterActive));
            OnPropertyChanged(nameof(IsEssaiDoneFilterActive));
        }

        private void RefreshEnCoursDashboard()
        {
            List<ProjetEnCoursSummary> dashboard = BuildProjetsEnCoursDashboard().ToList();

            projetsEnCoursDashboard = dashboard;
            projetsEnCoursDashboardCount = dashboard.Count;
            projetsEnAttentionDashboardCount = dashboard.Count(item => item.IsAttentionRequired);
            essaisTerminesDashboardTotal = dashboard.Sum(item => item.EssaisFaits);
            essaisEnCoursDashboardTotal = dashboard.Sum(item => item.EssaisEnCours);
            essaisRestantsDashboardTotal = dashboard.Sum(item => item.EssaisRestants);

            int essaisTotaux = dashboard.Sum(item => item.TotalEssais);
            avancementGlobalEnCoursPourcentage = essaisTotaux == 0
                ? 0
                : (int)Math.Round(essaisTerminesDashboardTotal * 100.0 / essaisTotaux);

            OnPropertyChanged(nameof(ProjetsEnCoursDashboard));
            OnPropertyChanged(nameof(HasProjetsEnCoursDashboard));
            OnPropertyChanged(nameof(ProjetsEnCoursDashboardCount));
            OnPropertyChanged(nameof(ProjetsEnAttentionDashboardCount));
            OnPropertyChanged(nameof(EssaisTerminesDashboardTotal));
            OnPropertyChanged(nameof(EssaisEnCoursDashboardTotal));
            OnPropertyChanged(nameof(EssaisRestantsDashboardTotal));
            OnPropertyChanged(nameof(AvancementGlobalEnCoursPourcentage));
            OnPropertyChanged(nameof(AvancementGlobalEnCoursLabel));
        }

        private IEnumerable<ProjetEnCoursSummary> BuildProjetsEnCoursDashboard()
        {
            return Projets
                .Where(IsProjetActifPourDashboard)
                .Select(BuildProjetEnCoursSummary)
                .OrderBy(item => item.StatusSortOrder)
                .ThenBy(item => item.DateDebutSortValue ?? DateTime.MaxValue)
                .ThenBy(item => item.NomProduit, StringComparer.CurrentCultureIgnoreCase)
                .ToList();
        }

        private bool IsProjetActifPourDashboard(Projet projet)
        {
            return !projet.EstArchive && NormalizeText(projet.Statut).Contains("cours");
        }

        private ProjetEnCoursSummary BuildProjetEnCoursSummary(Projet projet)
        {
            List<EssaiSuivi> essaisConcernes = projet.Essais
                .Where(essai => essai.EstConcerne)
                .ToList();

            List<string> essaisFaits = essaisConcernes
                .Where(IsEssaiDone)
                .Select(essai => essai.NomEssai ?? string.Empty)
                .Where(nom => !string.IsNullOrWhiteSpace(nom))
                .OrderBy(nom => nom, StringComparer.CurrentCultureIgnoreCase)
                .ToList();

            List<string> essaisEnCours = essaisConcernes
                .Where(IsEssaiInProgress)
                .Select(essai => essai.NomEssai ?? string.Empty)
                .Where(nom => !string.IsNullOrWhiteSpace(nom))
                .OrderBy(nom => nom, StringComparer.CurrentCultureIgnoreCase)
                .ToList();

            List<string> essaisAFaire = essaisConcernes
                .Where(IsEssaiToProcess)
                .Select(essai => essai.NomEssai ?? string.Empty)
                .Where(nom => !string.IsNullOrWhiteSpace(nom))
                .OrderBy(nom => nom, StringComparer.CurrentCultureIgnoreCase)
                .ToList();

            int totalEssais = essaisConcernes.Count;
            int doneCount = essaisFaits.Count;
            int inProgressCount = essaisEnCours.Count;
            int todoCount = essaisAFaire.Count;
            int remainingCount = Math.Max(0, totalEssais - doneCount);
            int progression = totalEssais == 0
                ? 0
                : (int)Math.Round(doneCount * 100.0 / totalEssais);
            bool isAttentionRequired = remainingCount > 0 && inProgressCount == 0;
            bool isNearlyDone = remainingCount <= 2 && totalEssais > 0;

            IReadOnlyList<string> donePreview = essaisFaits.Take(4).ToList();
            IReadOnlyList<string> remainingPreview = essaisEnCours
                .Concat(essaisAFaire)
                .Take(4)
                .ToList();

            return new ProjetEnCoursSummary(
                numeroProjet: projet.NumeroProjet ?? string.Empty,
                nomProduit: projet.NomProduit ?? string.Empty,
                client: projet.Client ?? "N/A",
                statut: projet.Statut ?? string.Empty,
                dateDebutLabel: projet.DateDebutValue?.ToString("dd/MM/yyyy", CultureInfo.GetCultureInfo("fr-FR")) ?? "Date à préciser",
                dateDebutSortValue: projet.DateDebutValue,
                totalEssais: totalEssais,
                essaisFaits: doneCount,
                essaisEnCours: inProgressCount,
                essaisAFaire: todoCount,
                essaisRestants: remainingCount,
                progressionPourcentage: progression,
                progressionLabel: $"{doneCount}/{totalEssais} essais faits",
                etatPilotage: BuildProjetEnCoursPilotageLabel(isAttentionRequired, isNearlyDone, inProgressCount),
                resumeEtat: $"{doneCount} faits · {inProgressCount} en cours · {remainingCount} restants",
                isAttentionRequired: isAttentionRequired,
                isNearlyDone: isNearlyDone,
                statusSortOrder: GetProjectStatusSortIndex(projet.Statut),
                essaisFaitsPreview: donePreview,
                essaisRestantsPreview: remainingPreview,
                essaisFaitsOverflowText: essaisFaits.Count > donePreview.Count ? $"+{essaisFaits.Count - donePreview.Count}" : string.Empty,
                essaisRestantsOverflowText: (essaisEnCours.Count + essaisAFaire.Count) > remainingPreview.Count ? $"+{(essaisEnCours.Count + essaisAFaire.Count) - remainingPreview.Count}" : string.Empty);
        }

        private static string BuildProjetEnCoursPilotageLabel(bool isAttentionRequired, bool isNearlyDone, int essaisEnCoursCount)
        {
            if (isAttentionRequired)
            {
                return "Attention";
            }

            if (isNearlyDone)
            {
                return "Finalisation";
            }

            if (essaisEnCoursCount > 0)
            {
                return "Cadencé";
            }

            return "Lancé";
        }

        private void UpdateProjectStatusFromEssais(Projet projet)
        {
            string statutProjet = NormalizeText(projet.Statut);
            bool isPreQualificationInProgress =
                statutProjet.Contains("pre") &&
                statutProjet.Contains("qualification") &&
                statutProjet.Contains("cours");
            bool isQualificationInProgress =
                !statutProjet.Contains("pre") &&
                statutProjet.Contains("qualification") &&
                statutProjet.Contains("cours");
            bool isReportInProgress =
                statutProjet.Contains("rapport") &&
                statutProjet.Contains("cours");

            if (!isPreQualificationInProgress && !isQualificationInProgress && !isReportInProgress)
            {
                return;
            }

            List<EssaiSuivi> essaisPreQualification = projet.EssaisPreQualification.ToList();
            bool allPreQualificationEssaisCompleted = essaisPreQualification.All(essai =>
                !essai.EstConcerne || IsEssaiDoneAndOk(essai));
            bool allProjectEssaisCompleted = projet.Essais.All(essai =>
                !essai.EstConcerne || IsEssaiDoneAndOk(essai));

            if (allProjectEssaisCompleted)
            {
                projet.Statut = "Rapport en cours";
            }
            else if (!allPreQualificationEssaisCompleted)
            {
                projet.Statut = "Pré-qualification en cours";
            }
            else
            {
                projet.Statut = "Qualification en cours";
            }
        }

        public bool MoveAgendaTaskToDay(AgendaTaskItem task, AgendaWorkDay targetDay, double? dropY = null)
        {
            if (!CanPlanOnAgendaDay(targetDay))
            {
                return false;
            }

            SaveAgendaUndoSnapshot();
            RemoveAgendaTaskFromAllBuckets(task);
            task.ScheduledStartMinutes = GetScheduledStartMinutes(targetDay, task, dropY);
            targetDay.PlannedTasks.Add(task);
            RecalculateAgendaWeek();
            RefreshAgendaCollections();
            return true;
        }

        public void MoveAgendaTaskToBacklog(AgendaTaskItem task)
        {
            SaveAgendaUndoSnapshot();
            RemoveAgendaTaskFromAllBuckets(task);
            ResetAgendaTaskLayout(task);
            AgendaBacklogTasks.Add(task);
            RecalculateAgendaWeek();
            RefreshAgendaCollections();
        }

        public bool PlanAgendaTaskAtFirstAvailableSlot(AgendaTaskItem task)
        {
            if (!AgendaBacklogTasks.Contains(task) &&
                !AgendaWeekDays.Any(day => day.PlannedTasks.Contains(task)))
            {
                return false;
            }

            AgendaWorkDay? targetDay = AgendaWeekDays.FirstOrDefault(CanPlanOnAgendaDay);
            if (targetDay is null)
            {
                return false;
            }

            SaveAgendaUndoSnapshot();
            RemoveAgendaTaskFromAllBuckets(task);
            task.ScheduledStartMinutes = GetAgendaDayDefaultStartMinutes(targetDay);
            targetDay.PlannedTasks.Add(task);
            RecalculateAgendaWeek();
            RefreshAgendaCollections();
            return true;
        }

        public void ClearAgendaSearch()
        {
            AgendaSearchText = string.Empty;
        }

        public void UndoAgendaLastAction()
        {
            if (agendaUndoSnapshots.Count == 0)
            {
                return;
            }

            AgendaUndoSnapshot snapshot = agendaUndoSnapshots.Pop();
            RestoreAgendaSnapshot(snapshot);
            OnPropertyChanged(nameof(CanUndoAgenda));
        }

        private ObservableCollection<AgendaWorkDay> CreateAgendaWeekDays()
        {
            var result = new ObservableCollection<AgendaWorkDay>();
            DateTime today = DateTime.Today;
            int offset = ((int)today.DayOfWeek + 6) % 7;
            DateTime monday = today.AddDays(-offset);
            string[] labels = { "Lundi", "Mardi", "Mercredi", "Jeudi", "Vendredi" };

            for (int i = 0; i < labels.Length; i++)
            {
                var day = new AgendaWorkDay(labels[i], monday.AddDays(i));
                day.PropertyChanged += AgendaWorkDay_PropertyChanged;
                day.PlannedTasks.CollectionChanged += AgendaPlannedTasks_CollectionChanged;
                result.Add(day);
            }

            int weekNumber = ISOWeek.GetWeekOfYear(monday);
            AgendaWeekTitle = $"Semaine en cours S{weekNumber:00} · {monday:dd/MM} - {monday.AddDays(4):dd/MM} · 1 j = 7 h";
            return result;
        }

        private void AgendaWorkDay_PropertyChanged(object? sender, PropertyChangedEventArgs e)
        {
            if (isRestoringAgendaState)
            {
                return;
            }

            if (sender is AgendaWorkDay &&
                e.PropertyName is nameof(AgendaWorkDay.IsWorkingDay) or
                    nameof(AgendaWorkDay.StartTimeText) or
                    nameof(AgendaWorkDay.EndTimeText) or
                    nameof(AgendaWorkDay.IsLunchBreakEnabled) or
                    nameof(AgendaWorkDay.LunchStartText) or
                    nameof(AgendaWorkDay.LunchEndText))
            {
                RecalculateAgendaWeek();
            }
        }

        private void AgendaPlannedTasks_CollectionChanged(object? sender, NotifyCollectionChangedEventArgs e)
        {
            if (isRestoringAgendaState)
            {
                return;
            }

            if (sender is not ObservableCollection<AgendaTaskItem> plannedTasks)
            {
                return;
            }

            AgendaWorkDay? day = AgendaWeekDays.FirstOrDefault(item => ReferenceEquals(item.PlannedTasks, plannedTasks));
            if (day != null)
            {
                RecalculateAgendaWeek();
                RefreshAgendaCollections();
            }
        }

        private IEnumerable<AgendaTaskItem> GetUnplannedAgendaTasks()
        {
            HashSet<string> plannedKeys = AgendaWeekDays
                .SelectMany(day => day.PlannedTasks)
                .Select(task => task.TaskKey)
                .ToHashSet(StringComparer.OrdinalIgnoreCase);

            return AgendaBacklogTasks.Where(task => !plannedKeys.Contains(task.TaskKey));
        }

        private bool MatchesAgendaSearch(AgendaTaskItem task)
        {
            if (string.IsNullOrWhiteSpace(AgendaSearchText))
            {
                return true;
            }

            string searchValue = NormalizeText(AgendaSearchText);

            return NormalizeText(task.NomEssai).Contains(searchValue) ||
                   NormalizeText(task.NomProduit).Contains(searchValue) ||
                   NormalizeText(task.NumeroProjet).Contains(searchValue);
        }

        private void RefreshAgendaTasks()
        {
            Dictionary<string, AgendaTaskItem> existingTasks = AgendaBacklogTasks
                .Concat(AgendaWeekDays.SelectMany(day => day.PlannedTasks))
                .ToDictionary(task => task.TaskKey, StringComparer.OrdinalIgnoreCase);

            HashSet<string> validKeys = new(StringComparer.OrdinalIgnoreCase);

            foreach (Projet projet in Projets)
            {
                if (projet.EstArchive)
                {
                    continue;
                }

                foreach (EssaiSuivi essai in projet.Essais.Where(ShouldAppearInAgenda))
                {
                    string key = BuildAgendaTaskKey(projet, essai);
                    validKeys.Add(key);

                    if (!existingTasks.TryGetValue(key, out AgendaTaskItem? task))
                    {
                        task = new AgendaTaskItem { TaskKey = key };
                        AgendaBacklogTasks.Add(task);
                    }

                    double dureeJours = AgendaDureesEssais.TryGetValue(essai.NomEssai ?? string.Empty, out double value)
                        ? value
                        : 0.25;

                    task.NumeroProjet = projet.NumeroProjet ?? string.Empty;
                    task.NomProduit = projet.NomProduit ?? string.Empty;
                    task.NomEssai = essai.NomEssai ?? string.Empty;
                    task.DureeJours = dureeJours;
                    task.DureeHeures = dureeJours * AgendaHoursPerDay;
                }
            }

            foreach (AgendaTaskItem obsolete in existingTasks.Values.Where(task => !validKeys.Contains(task.TaskKey)).ToList())
            {
                RemoveAgendaTaskFromAllBuckets(obsolete);
            }

            EnsureAgendaBacklogConsistency();

            RecalculateAgendaWeek();

            RefreshAgendaCollections();
        }

        private void SaveAgendaUndoSnapshot()
        {
            AgendaUndoSnapshot snapshot = CaptureAgendaSnapshot();
            agendaUndoSnapshots.Push(snapshot);

            while (agendaUndoSnapshots.Count > 30)
            {
                AgendaUndoSnapshot[] preserved = agendaUndoSnapshots.Reverse().Take(30).ToArray();
                agendaUndoSnapshots.Clear();

                for (int i = preserved.Length - 1; i >= 0; i--)
                {
                    agendaUndoSnapshots.Push(preserved[i]);
                }
            }

            OnPropertyChanged(nameof(CanUndoAgenda));
        }

        private AgendaUndoSnapshot CaptureAgendaSnapshot()
        {
            var placements = new List<AgendaTaskPlacementSnapshot>();

            foreach (AgendaTaskItem task in AgendaBacklogTasks)
            {
                placements.Add(new AgendaTaskPlacementSnapshot(task.TaskKey, null, null, null));
            }

            for (int dayIndex = 0; dayIndex < AgendaWeekDays.Count; dayIndex++)
            {
                AgendaWorkDay day = AgendaWeekDays[dayIndex];
                for (int orderIndex = 0; orderIndex < day.PlannedTasks.Count; orderIndex++)
                {
                    AgendaTaskItem task = day.PlannedTasks[orderIndex];
                    placements.Add(new AgendaTaskPlacementSnapshot(task.TaskKey, dayIndex, task.ScheduledStartMinutes, orderIndex));
                }
            }

            return new AgendaUndoSnapshot(placements);
        }

        private void RestoreAgendaSnapshot(AgendaUndoSnapshot snapshot)
        {
            Dictionary<string, AgendaTaskItem> tasksByKey = AgendaBacklogTasks
                .Concat(AgendaWeekDays.SelectMany(day => day.PlannedTasks))
                .ToDictionary(task => task.TaskKey, StringComparer.OrdinalIgnoreCase);

            isRestoringAgendaState = true;

            try
            {
                AgendaBacklogTasks.Clear();

                foreach (AgendaWorkDay day in AgendaWeekDays)
                {
                    day.PlannedTasks.Clear();
                }

                foreach (AgendaTaskPlacementSnapshot placement in snapshot.Placements
                             .OrderBy(item => item.DayIndex ?? int.MaxValue)
                             .ThenBy(item => item.OrderIndex ?? int.MaxValue))
                {
                    if (!tasksByKey.TryGetValue(placement.TaskKey, out AgendaTaskItem? task))
                    {
                        continue;
                    }

                    if (placement.DayIndex.HasValue &&
                        placement.DayIndex.Value >= 0 &&
                        placement.DayIndex.Value < AgendaWeekDays.Count)
                    {
                        task.ScheduledStartMinutes = placement.ScheduledStartMinutes;
                        AgendaWeekDays[placement.DayIndex.Value].PlannedTasks.Add(task);
                    }
                    else
                    {
                        ResetAgendaTaskLayout(task);
                        AgendaBacklogTasks.Add(task);
                    }
                }
            }
            finally
            {
                isRestoringAgendaState = false;
            }

            RecalculateAgendaWeek();
            RefreshAgendaCollections();
        }

        private bool SaveEssaisUndoSnapshotIfNeeded(Projet projet, Func<EssaiSuivi, bool> hasMeaningfulChange)
        {
            if (!projet.Essais.Any(hasMeaningfulChange))
            {
                return false;
            }

            EssaisUndoSnapshot snapshot = CaptureEssaisSnapshot(projet);
            essaisUndoSnapshots.Push(snapshot);

            while (essaisUndoSnapshots.Count > 30)
            {
                EssaisUndoSnapshot[] preserved = essaisUndoSnapshots.Reverse().Take(30).ToArray();
                essaisUndoSnapshots.Clear();

                for (int i = preserved.Length - 1; i >= 0; i--)
                {
                    essaisUndoSnapshots.Push(preserved[i]);
                }
            }

            OnPropertyChanged(nameof(CanUndoEssaisBulkAction));
            return true;
        }

        private EssaisUndoSnapshot CaptureEssaisSnapshot(Projet projet)
        {
            var states = projet.Essais
                .Select((essai, index) => new EssaiStateSnapshot(index, essai.Statut, essai.ResultatTraitement))
                .ToList();

            return new EssaisUndoSnapshot(projet, states);
        }

        private void RestoreEssaisSnapshot(EssaisUndoSnapshot snapshot)
        {
            if (!Projets.Contains(snapshot.Projet))
            {
                return;
            }

            SelectedProjetEssais = snapshot.Projet;

            foreach (EssaiStateSnapshot state in snapshot.States)
            {
                if (state.Index < 0 || state.Index >= snapshot.Projet.Essais.Count)
                {
                    continue;
                }

                EssaiSuivi essai = snapshot.Projet.Essais[state.Index];
                essai.Statut = state.Statut;
                essai.ResultatTraitement = state.ResultatTraitement;
            }

            RefreshEssaiCollections();
            RefreshSelectedProjectStatistics();
            RefreshAgendaTasks();
        }

        private ProjectTableUndoSnapshot CaptureProjectTableSnapshot()
        {
            var states = Projets
                .Select(projet => new ProjectTableStateSnapshot(
                    projet,
                    projet.NumeroProjet,
                    projet.NomProduit,
                    projet.FamilleProduit,
                    projet.ReferenceProduit,
                    projet.Client,
                    projet.Demandeur,
                    projet.TypeActivite,
                    projet.DossierRacine,
                    projet.Statut,
                    projet.EstArchive,
                    projet.DateDebut,
                    projet.DatePrevisionnelle,
                    projet.DateFin,
                    projet.Commentaires))
                .ToList();

            return new ProjectTableUndoSnapshot(states);
        }

        private void RestoreProjectTableSnapshot(ProjectTableUndoSnapshot snapshot)
        {
            foreach (ProjectTableStateSnapshot state in snapshot.States)
            {
                if (!Projets.Contains(state.Projet))
                {
                    continue;
                }

                state.Projet.NumeroProjet = state.NumeroProjet;
                state.Projet.NomProduit = state.NomProduit;
                state.Projet.FamilleProduit = state.FamilleProduit;
                state.Projet.ReferenceProduit = state.ReferenceProduit;
                state.Projet.Client = state.Client;
                state.Projet.Demandeur = state.Demandeur;
                state.Projet.TypeActivite = state.TypeActivite;
                state.Projet.DossierRacine = state.DossierRacine;
                state.Projet.Statut = state.Statut;
                state.Projet.EstArchive = state.EstArchive;
                state.Projet.DateDebut = state.DateDebut;
                state.Projet.DatePrevisionnelle = state.DatePrevisionnelle;
                state.Projet.DateFin = state.DateFin;
                state.Projet.Commentaires = state.Commentaires;
            }

            ProjetsView.Refresh();
            ProjetsEssaisView.Refresh();
            RefreshStatistics();
            RefreshSelectedProjectStatistics();
            RefreshAgendaTasks();
        }

        private static bool ShouldAppearInAgenda(EssaiSuivi essai)
        {
            return essai.EstConcerne && !IsEssaiDone(essai);
        }

        private static string BuildAgendaTaskKey(Projet projet, EssaiSuivi essai)
        {
            return $"{projet.NumeroProjet}|{projet.NomProduit}|{essai.NomEssai}";
        }

        private void RemoveAgendaTaskFromAllBuckets(AgendaTaskItem task)
        {
            foreach (AgendaTaskItem backlogTask in AgendaBacklogTasks
                         .Where(item => IsSameAgendaTask(item, task))
                         .ToList())
            {
                AgendaBacklogTasks.Remove(backlogTask);
            }

            foreach (AgendaWorkDay day in AgendaWeekDays)
            {
                foreach (AgendaTaskItem plannedTask in day.PlannedTasks
                             .Where(item => IsSameAgendaTask(item, task))
                             .ToList())
                {
                    day.PlannedTasks.Remove(plannedTask);
                }
            }
        }

        private static bool IsSameAgendaTask(AgendaTaskItem left, AgendaTaskItem right)
        {
            return ReferenceEquals(left, right) ||
                   string.Equals(left.TaskKey, right.TaskKey, StringComparison.OrdinalIgnoreCase);
        }

        private void EnsureAgendaBacklogConsistency()
        {
            HashSet<string> plannedKeys = AgendaWeekDays
                .SelectMany(day => day.PlannedTasks)
                .Select(task => task.TaskKey)
                .ToHashSet(StringComparer.OrdinalIgnoreCase);

            foreach (AgendaTaskItem duplicate in AgendaBacklogTasks
                         .Where(task => plannedKeys.Contains(task.TaskKey))
                         .ToList())
            {
                AgendaBacklogTasks.Remove(duplicate);
            }
        }

        private void RecalculateAgendaWeek()
        {
            var occupiedRangesByDay = new Dictionary<AgendaWorkDay, List<AgendaTimeRange>>();

            foreach (AgendaWorkDay day in AgendaWeekDays)
            {
                day.DisplaySegments.Clear();
                day.NonWorkingTimeBlocks.Clear();
                occupiedRangesByDay[day] = new List<AgendaTimeRange>();

                if (!TryGetAgendaDayLayout(
                        day,
                        out TimeSpan startTime,
                        out TimeSpan endTime,
                        out TimeSpan lunchStart,
                        out TimeSpan lunchEnd,
                        out TimeSpan displayStart,
                        out TimeSpan displayEnd))
                {
                    PopulateHourSlots(day, TimeSpan.FromHours(8), TimeSpan.FromHours(16));
                    AddAgendaNonWorkingBlock(day, AgendaDisplayStartTime, AgendaDisplayStartTime, AgendaDisplayEndTime, "Horaires invalides");
                    continue;
                }

                PopulateHourSlots(day, displayStart, displayEnd);
                UpdateAgendaNonWorkingBlocks(day, startTime, endTime, lunchStart, lunchEnd, displayStart, displayEnd);
            }

            for (int dayIndex = 0; dayIndex < AgendaWeekDays.Count; dayIndex++)
            {
                AgendaWorkDay anchorDay = AgendaWeekDays[dayIndex];

                foreach (AgendaTaskItem task in anchorDay.PlannedTasks
                             .Select((item, orderIndex) => new { Task = item, OrderIndex = orderIndex })
                             .OrderBy(item => item.Task.ScheduledStartMinutes ?? int.MaxValue)
                             .ThenBy(item => item.OrderIndex)
                             .Select(item => item.Task))
                {
                    PlaceTaskAcrossWeek(dayIndex, task, occupiedRangesByDay);
                }
            }
        }

        private void PlaceTaskAcrossWeek(
            int startDayIndex,
            AgendaTaskItem task,
            IDictionary<AgendaWorkDay, List<AgendaTimeRange>> occupiedRangesByDay)
        {
            double remainingHours = task.DureeHeures;
            bool firstSegment = true;
            AgendaTaskSegment? lastSegment = null;

            ResetAgendaTaskVisualLayout(task);

            for (int currentDayIndex = startDayIndex; currentDayIndex < AgendaWeekDays.Count && remainingHours > AgendaScheduleEpsilon; currentDayIndex++)
            {
                AgendaWorkDay currentDay = AgendaWeekDays[currentDayIndex];

                if (!TryGetAgendaDayLayout(currentDay, out TimeSpan startTime, out TimeSpan endTime, out TimeSpan lunchStart, out TimeSpan lunchEnd, out TimeSpan displayStart, out _))
                {
                    continue;
                }

                if (!currentDay.IsWorkingDay)
                {
                    continue;
                }

                TimeSpan searchStart = firstSegment && currentDayIndex == startDayIndex && task.ScheduledStartMinutes.HasValue
                    ? TimeSpan.FromMinutes(task.ScheduledStartMinutes.Value)
                    : startTime;

                foreach (AgendaTimeRange availableRange in GetAvailableAgendaRanges(
                             startTime,
                             endTime,
                             lunchStart,
                             lunchEnd,
                             searchStart,
                             occupiedRangesByDay[currentDay]))
                {
                    double availableHours = (availableRange.End - availableRange.Start).TotalHours;
                    if (availableHours <= AgendaScheduleEpsilon)
                    {
                        continue;
                    }

                    double segmentHours = Math.Min(remainingHours, availableHours);
                    TimeSpan segmentStart = availableRange.Start;
                    TimeSpan segmentEnd = segmentStart.Add(TimeSpan.FromHours(segmentHours));
                    double timelineTop = Math.Max(0, (segmentStart - displayStart).TotalHours * AgendaHourSlotHeight);
                    double blockHeight = Math.Max(30, ((segmentEnd - segmentStart).TotalHours * AgendaHourSlotHeight) - 4);

                    lastSegment = new AgendaTaskSegment
                    {
                        SourceTask = task,
                        NomEssai = firstSegment ? task.NomEssai : $"{task.NomEssai} (suite)",
                        NumeroProjet = task.NumeroProjet,
                        NomProduit = task.NomProduit,
                        DureeLabel = AgendaDurationFormatter.Format(segmentHours, segmentHours / AgendaHoursPerDay),
                        TimeRangeLabel = $"{segmentStart:hh\\:mm} - {segmentEnd:hh\\:mm}",
                        IsOverflow = false,
                        BlockHeight = blockHeight,
                        TimelineMargin = new Thickness(10, timelineTop, 10, 0),
                        IsContinuation = !firstSegment
                    };

                    currentDay.DisplaySegments.Add(lastSegment);
                    AddAgendaOccupiedRange(occupiedRangesByDay[currentDay], new AgendaTimeRange(segmentStart, segmentEnd));

                    if (firstSegment)
                    {
                        task.TimeRangeLabel = lastSegment.TimeRangeLabel;
                        task.IsOverflow = false;
                        task.TimelineTop = timelineTop;
                        task.BlockHeight = blockHeight;
                        task.TimelineMargin = lastSegment.TimelineMargin;
                    }

                    remainingHours -= segmentHours;
                    firstSegment = false;

                    if (remainingHours <= AgendaScheduleEpsilon)
                    {
                        break;
                    }
                }
            }

            if (remainingHours > AgendaScheduleEpsilon)
            {
                if (lastSegment != null)
                {
                    lastSegment.IsOverflow = true;
                }

                task.IsOverflow = true;
            }
            else if (firstSegment)
            {
                task.TimeRangeLabel = "Journée non travaillée";
                task.IsOverflow = true;
                task.TimelineTop = 0;
                task.BlockHeight = 52;
                task.TimelineMargin = new Thickness(10, 0, 10, 0);
            }
        }

        private static void ResetAgendaTaskVisualLayout(AgendaTaskItem task)
        {
            task.TimeRangeLabel = null;
            task.IsOverflow = false;
            task.TimelineTop = 0;
            task.BlockHeight = 52;
            task.TimelineMargin = new Thickness(10, 0, 10, 0);
        }

        private bool TryGetAgendaDayLayout(
            AgendaWorkDay day,
            out TimeSpan startTime,
            out TimeSpan endTime,
            out TimeSpan lunchStart,
            out TimeSpan lunchEnd,
            out TimeSpan displayStart,
            out TimeSpan displayEnd)
        {
            displayStart = TimeSpan.FromHours(8);
            displayEnd = TimeSpan.FromHours(16);
            lunchStart = TimeSpan.Zero;
            lunchEnd = TimeSpan.Zero;

            if (!TryParseTime(day.StartTimeText, out startTime) ||
                !TryParseTime(day.EndTimeText, out endTime) ||
                endTime <= startTime)
            {
                startTime = TimeSpan.FromHours(8);
                endTime = TimeSpan.FromHours(16);
                return false;
            }

            bool hasLunchBreak =
                day.IsLunchBreakEnabled &&
                TryParseTime(day.LunchStartText, out lunchStart) &&
                TryParseTime(day.LunchEndText, out lunchEnd) &&
                lunchEnd > lunchStart &&
                lunchStart >= startTime &&
                lunchEnd <= endTime;

            if (!hasLunchBreak)
            {
                lunchStart = TimeSpan.Zero;
                lunchEnd = TimeSpan.Zero;
            }

            displayStart = AgendaDisplayStartTime;
            displayEnd = AgendaDisplayEndTime;

            return true;
        }

        private static IEnumerable<AgendaTimeRange> GetAvailableAgendaRanges(
            TimeSpan startTime,
            TimeSpan endTime,
            TimeSpan lunchStart,
            TimeSpan lunchEnd,
            TimeSpan requestedStart,
            IEnumerable<AgendaTimeRange> occupiedRanges)
        {
            foreach (AgendaTimeRange workingRange in GetAgendaWorkingRanges(startTime, endTime, lunchStart, lunchEnd))
            {
                TimeSpan currentStart = MaxTime(workingRange.Start, requestedStart);
                if (currentStart >= workingRange.End)
                {
                    continue;
                }

                foreach (AgendaTimeRange occupiedRange in occupiedRanges
                             .Where(range => range.End > currentStart && range.Start < workingRange.End)
                             .OrderBy(range => range.Start))
                {
                    if (occupiedRange.Start > currentStart)
                    {
                        yield return new AgendaTimeRange(currentStart, MinTime(occupiedRange.Start, workingRange.End));
                    }

                    if (occupiedRange.End > currentStart)
                    {
                        currentStart = MinTime(MaxTime(currentStart, occupiedRange.End), workingRange.End);
                    }

                    if (currentStart >= workingRange.End)
                    {
                        break;
                    }
                }

                if (currentStart < workingRange.End)
                {
                    yield return new AgendaTimeRange(currentStart, workingRange.End);
                }
            }
        }

        private static IEnumerable<AgendaTimeRange> GetAgendaWorkingRanges(
            TimeSpan startTime,
            TimeSpan endTime,
            TimeSpan lunchStart,
            TimeSpan lunchEnd)
        {
            if (endTime <= startTime)
            {
                yield break;
            }

            bool hasLunchBreak = lunchEnd > lunchStart;

            if (!hasLunchBreak)
            {
                yield return new AgendaTimeRange(startTime, endTime);
                yield break;
            }

            if (lunchStart > startTime)
            {
                yield return new AgendaTimeRange(startTime, lunchStart);
            }

            if (lunchEnd < endTime)
            {
                yield return new AgendaTimeRange(lunchEnd, endTime);
            }
        }

        private int? GetScheduledStartMinutes(AgendaWorkDay day, AgendaTaskItem task, double? dropY)
        {
            if (!dropY.HasValue ||
                !TryGetAgendaDayLayout(day, out TimeSpan startTime, out TimeSpan endTime, out TimeSpan lunchStart, out TimeSpan lunchEnd, out TimeSpan displayStart, out TimeSpan displayEnd))
            {
                return task.ScheduledStartMinutes;
            }

            double totalDisplayHours = Math.Max(1, (displayEnd - displayStart).TotalHours);

            double limitedY = Math.Max(0, Math.Min(dropY.Value, totalDisplayHours * AgendaHourSlotHeight));
            double rawMinutes = displayStart.TotalMinutes + (limitedY / AgendaHourSlotHeight * 60);
            int snappedMinutes = (int)(Math.Round(rawMinutes / AgendaSnapMinutes) * AgendaSnapMinutes);

            TimeSpan candidate = TimeSpan.FromMinutes(snappedMinutes);
            TimeSpan adjustedCandidate = ClampAgendaStartToWorkingRange(candidate, startTime, endTime, lunchStart, lunchEnd);

            return (int)adjustedCandidate.TotalMinutes;
        }

        private bool CanPlanOnAgendaDay(AgendaWorkDay day)
        {
            if (!day.IsWorkingDay ||
                !TryGetAgendaDayLayout(day, out TimeSpan startTime, out TimeSpan endTime, out TimeSpan lunchStart, out TimeSpan lunchEnd, out _, out _))
            {
                return false;
            }

            return GetAgendaWorkingRanges(startTime, endTime, lunchStart, lunchEnd).Any();
        }

        private int? GetAgendaDayDefaultStartMinutes(AgendaWorkDay day)
        {
            if (!TryGetAgendaDayLayout(day, out TimeSpan startTime, out TimeSpan endTime, out TimeSpan lunchStart, out TimeSpan lunchEnd, out _, out _))
            {
                return null;
            }

            return (int)(GetAgendaWorkingRanges(startTime, endTime, lunchStart, lunchEnd).FirstOrDefault().Start.TotalMinutes);
        }

        private static TimeSpan ClampAgendaStartToWorkingRange(
            TimeSpan candidate,
            TimeSpan startTime,
            TimeSpan endTime,
            TimeSpan lunchStart,
            TimeSpan lunchEnd)
        {
            List<AgendaTimeRange> workingRanges = GetAgendaWorkingRanges(startTime, endTime, lunchStart, lunchEnd).ToList();
            if (workingRanges.Count == 0)
            {
                return startTime;
            }

            foreach (AgendaTimeRange workingRange in workingRanges)
            {
                if (candidate <= workingRange.Start)
                {
                    return workingRange.Start;
                }

                if (candidate < workingRange.End)
                {
                    return candidate;
                }
            }

            AgendaTimeRange lastRange = workingRanges[^1];
            TimeSpan latestStart = lastRange.End.Subtract(TimeSpan.FromMinutes(AgendaSnapMinutes));
            return latestStart > lastRange.Start ? latestStart : lastRange.Start;
        }

        private static void AddAgendaOccupiedRange(List<AgendaTimeRange> ranges, AgendaTimeRange newRange)
        {
            if (newRange.End <= newRange.Start)
            {
                return;
            }

            ranges.Add(newRange);
        }

        private static void UpdateAgendaNonWorkingBlocks(
            AgendaWorkDay day,
            TimeSpan startTime,
            TimeSpan endTime,
            TimeSpan lunchStart,
            TimeSpan lunchEnd,
            TimeSpan displayStart,
            TimeSpan displayEnd)
        {
            if (!day.IsWorkingDay)
            {
                AddAgendaNonWorkingBlock(day, displayStart, displayStart, displayEnd, string.Empty);
                return;
            }

            AddAgendaNonWorkingBlock(day, displayStart, displayStart, startTime, string.Empty);

            if (lunchEnd > lunchStart)
            {
                AddAgendaNonWorkingBlock(day, displayStart, lunchStart, lunchEnd, $"Pause déjeuner {lunchStart:hh\\:mm} - {lunchEnd:hh\\:mm}");
            }

            AddAgendaNonWorkingBlock(day, displayStart, endTime, displayEnd, string.Empty);
        }

        private static void AddAgendaNonWorkingBlock(
            AgendaWorkDay day,
            TimeSpan displayStart,
            TimeSpan rangeStart,
            TimeSpan rangeEnd,
            string label)
        {
            TimeSpan blockStart = MaxTime(rangeStart, AgendaDisplayStartTime);
            TimeSpan blockEnd = MinTime(rangeEnd, AgendaDisplayEndTime);

            if (blockEnd <= blockStart)
            {
                return;
            }

            double timelineTop = Math.Max(0, (blockStart - displayStart).TotalHours * AgendaHourSlotHeight);
            double blockHeight = Math.Max(8, ((blockEnd - blockStart).TotalHours * AgendaHourSlotHeight) - 1);

            day.NonWorkingTimeBlocks.Add(new AgendaTimeBlock
            {
                Label = label,
                BlockHeight = blockHeight,
                TimelineMargin = new Thickness(0, timelineTop, 0, 0)
            });
        }

        private static TimeSpan MinTime(TimeSpan left, TimeSpan right)
        {
            return left <= right ? left : right;
        }

        private static TimeSpan MaxTime(TimeSpan left, TimeSpan right)
        {
            return left >= right ? left : right;
        }

        private static bool TryParseTime(string? value, out TimeSpan time)
        {
            return TimeSpan.TryParse(value, CultureInfo.InvariantCulture, out time) ||
                   TimeSpan.TryParse(value, CultureInfo.GetCultureInfo("fr-FR"), out time);
        }

        private static void PopulateHourSlots(AgendaWorkDay day, TimeSpan start, TimeSpan end)
        {
            day.HourSlots.Clear();

            start = AgendaDisplayStartTime;
            end = AgendaDisplayEndTime;

            TimeSpan current = start;
            while (current < end)
            {
                day.HourSlots.Add($"{current:hh\\:mm}");
                current = current.Add(TimeSpan.FromHours(1));
            }

            day.TimelineHeight = day.HourSlots.Count * AgendaHourSlotHeight;
            day.EndHourLabel = $"{end:hh\\:mm}";
        }

        private static void ResetAgendaTaskLayout(AgendaTaskItem task)
        {
            task.TimeRangeLabel = null;
            task.IsOverflow = false;
            task.TimelineTop = 0;
            task.BlockHeight = 52;
            task.ScheduledStartMinutes = null;
            task.TimelineMargin = new Thickness(10, 0, 10, 0);
        }

        private void RefreshAgendaCollections()
        {
            OnPropertyChanged(nameof(AgendaBacklogVisibleTasks));
            OnPropertyChanged(nameof(AgendaBacklogVisibleCount));
            OnPropertyChanged(nameof(AgendaBacklogRemainingCount));
            OnPropertyChanged(nameof(HasAgendaSearch));
            OnPropertyChanged(nameof(HasHiddenAgendaBacklogTasks));
            OnPropertyChanged(nameof(AgendaBacklogFilterSummary));
            OnPropertyChanged(nameof(AgendaPlannedTaskCount));
            OnPropertyChanged(nameof(AgendaOverflowTaskCount));
            OnPropertyChanged(nameof(AgendaPlannedDurationLabel));
            OnPropertyChanged(nameof(AgendaDisplayHourMarkers));
            OnPropertyChanged(nameof(AgendaDisplayGridSlots));
            OnPropertyChanged(nameof(AgendaTimelineHeight));
            OnPropertyChanged(nameof(AgendaWeekTitle));
            OnPropertyChanged(nameof(CanUndoAgenda));
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

        private static bool RepairLoadedProject(Projet projet)
        {
            bool hasRepairs = false;

            hasRepairs |= RepairStringProperty(projet.NumeroProjet, repaired => projet.NumeroProjet = repaired);
            hasRepairs |= RepairStringProperty(projet.NomProduit, repaired => projet.NomProduit = repaired);
            hasRepairs |= RepairStringProperty(projet.FamilleProduit, repaired => projet.FamilleProduit = repaired);
            hasRepairs |= RepairStringProperty(projet.ReferenceProduit, repaired => projet.ReferenceProduit = repaired);
            hasRepairs |= RepairStringProperty(projet.Client, repaired => projet.Client = repaired);
            hasRepairs |= RepairStringProperty(projet.Demandeur, repaired => projet.Demandeur = repaired);
            hasRepairs |= RepairStringProperty(projet.TypeActivite, repaired => projet.TypeActivite = repaired);
            hasRepairs |= RepairStringProperty(projet.DossierRacine, repaired => projet.DossierRacine = repaired);
            hasRepairs |= RepairStringProperty(projet.Statut, repaired => projet.Statut = repaired);
            hasRepairs |= RepairStringProperty(projet.DateDebut, repaired => projet.DateDebut = repaired);
            hasRepairs |= RepairStringProperty(projet.DatePrevisionnelle, repaired => projet.DatePrevisionnelle = repaired);
            hasRepairs |= RepairStringProperty(projet.DateFin, repaired => projet.DateFin = repaired);
            hasRepairs |= RepairStringProperty(projet.Commentaires, repaired => projet.Commentaires = repaired);

            foreach (EssaiSuivi essai in projet.Essais)
            {
                hasRepairs |= RepairStringProperty(essai.NomEssai, repaired => essai.NomEssai = repaired);
                hasRepairs |= RepairStringProperty(essai.Statut, repaired => essai.Statut = repaired);
                hasRepairs |= RepairStringProperty(essai.ResultatTraitement, repaired => essai.ResultatTraitement = repaired);
            }

            return hasRepairs;
        }

        private static bool RepairStringProperty(string? currentValue, Action<string> assign)
        {
            string repairedValue = RepairMojibakeIfNeeded(currentValue);
            if (string.Equals(currentValue, repairedValue, StringComparison.Ordinal))
            {
                return false;
            }

            assign(repairedValue);
            return true;
        }

        private static string RepairMojibakeIfNeeded(string? value)
        {
            string input = value ?? string.Empty;
            if (!input.Contains('\u00C3') && !input.Contains('\u00C2') && !input.Contains('\u00E2'))
            {
                return input;
            }

            try
            {
                return Encoding.UTF8.GetString(Encoding.Latin1.GetBytes(input));
            }
            catch (ArgumentException)
            {
                return input;
            }
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

        public sealed class ProjetEnCoursSummary
        {
            public ProjetEnCoursSummary(
                string numeroProjet,
                string nomProduit,
                string client,
                string statut,
                string dateDebutLabel,
                DateTime? dateDebutSortValue,
                int totalEssais,
                int essaisFaits,
                int essaisEnCours,
                int essaisAFaire,
                int essaisRestants,
                int progressionPourcentage,
                string progressionLabel,
                string etatPilotage,
                string resumeEtat,
                bool isAttentionRequired,
                bool isNearlyDone,
                int statusSortOrder,
                IReadOnlyList<string> essaisFaitsPreview,
                IReadOnlyList<string> essaisRestantsPreview,
                string essaisFaitsOverflowText,
                string essaisRestantsOverflowText)
            {
                NumeroProjet = numeroProjet;
                NomProduit = nomProduit;
                Client = client;
                Statut = statut;
                DateDebutLabel = dateDebutLabel;
                DateDebutSortValue = dateDebutSortValue;
                TotalEssais = totalEssais;
                EssaisFaits = essaisFaits;
                EssaisEnCours = essaisEnCours;
                EssaisAFaire = essaisAFaire;
                EssaisRestants = essaisRestants;
                ProgressionPourcentage = progressionPourcentage;
                ProgressionLabel = progressionLabel;
                EtatPilotage = etatPilotage;
                ResumeEtat = resumeEtat;
                IsAttentionRequired = isAttentionRequired;
                IsNearlyDone = isNearlyDone;
                StatusSortOrder = statusSortOrder;
                EssaisFaitsPreview = essaisFaitsPreview;
                EssaisRestantsPreview = essaisRestantsPreview;
                EssaisFaitsOverflowText = essaisFaitsOverflowText;
                EssaisRestantsOverflowText = essaisRestantsOverflowText;
            }

            public string NumeroProjet { get; }

            public string NomProduit { get; }

            public string Client { get; }

            public string Statut { get; }

            public string DateDebutLabel { get; }

            public DateTime? DateDebutSortValue { get; }

            public int TotalEssais { get; }

            public int EssaisFaits { get; }

            public int EssaisEnCours { get; }

            public int EssaisAFaire { get; }

            public int EssaisRestants { get; }

            public int ProgressionPourcentage { get; }

            public string ProgressionLabel { get; }

            public string EtatPilotage { get; }

            public string ResumeEtat { get; }

            public bool IsAttentionRequired { get; }

            public bool IsNearlyDone { get; }

            public int StatusSortOrder { get; }

            public IReadOnlyList<string> EssaisFaitsPreview { get; }

            public IReadOnlyList<string> EssaisRestantsPreview { get; }

            public string EssaisFaitsOverflowText { get; }

            public string EssaisRestantsOverflowText { get; }

            public bool HasEssaisFaits => EssaisFaitsPreview.Count > 0;

            public bool HasEssaisRestants => EssaisRestantsPreview.Count > 0;

            public bool HasEssaisFaitsOverflow => !string.IsNullOrWhiteSpace(EssaisFaitsOverflowText);

            public bool HasEssaisRestantsOverflow => !string.IsNullOrWhiteSpace(EssaisRestantsOverflowText);
        }

        private sealed class DelegateComparer : IComparer
        {
            private readonly Comparison<object?> comparison;

            public DelegateComparer(Comparison<object?> comparison)
            {
                this.comparison = comparison;
            }

            public int Compare(object? x, object? y)
            {
                return comparison(x, y);
            }
        }

        private readonly record struct AgendaTimeRange(TimeSpan Start, TimeSpan End);

        private sealed record AgendaTaskPlacementSnapshot(string TaskKey, int? DayIndex, int? ScheduledStartMinutes, int? OrderIndex);

        private sealed record AgendaUndoSnapshot(IReadOnlyList<AgendaTaskPlacementSnapshot> Placements);

        private sealed record EssaiStateSnapshot(int Index, string? Statut, string? ResultatTraitement);

        private sealed record EssaisUndoSnapshot(Projet Projet, IReadOnlyList<EssaiStateSnapshot> States);

        private sealed record ProjectTableStateSnapshot(
            Projet Projet,
            string? NumeroProjet,
            string? NomProduit,
            string? FamilleProduit,
            string? ReferenceProduit,
            string? Client,
            string? Demandeur,
            string? TypeActivite,
            string? DossierRacine,
            string? Statut,
            bool EstArchive,
            string? DateDebut,
            string? DatePrevisionnelle,
            string? DateFin,
            string? Commentaires);

        private sealed record ProjectTableUndoSnapshot(IReadOnlyList<ProjectTableStateSnapshot> States);

        private sealed class ReferencePropertiesSettings
        {
            public List<string>? Clients { get; set; }

            public List<string>? Demandeurs { get; set; }

            public List<string>? FamillesProduit { get; set; }

            public List<string>? TypesActivite { get; set; }

            public Dictionary<string, List<string>>? EssaisParFamille { get; set; }
        }

        public sealed class EssaiDefaultSelectionItem : INotifyPropertyChanged
        {
            private bool isSelected;

            public EssaiDefaultSelectionItem(string nomEssai, bool isSelected)
            {
                NomEssai = nomEssai;
                this.isSelected = isSelected;
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

        public sealed class AgendaHourMarker
        {
            public AgendaHourMarker(string label, double offset)
            {
                Label = label;
                Offset = offset;
            }

            public string Label { get; }

            public double Offset { get; }
        }
    }
}
