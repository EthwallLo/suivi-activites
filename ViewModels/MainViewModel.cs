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
using System.Windows;
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
        private const double AgendaHoursPerDay = 8.0;
        private const double AgendaHourSlotHeight = 48.0;
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
        private string agendaWeekTitle = string.Empty;
        private readonly Stack<AgendaUndoSnapshot> agendaUndoSnapshots = new();
        private readonly Stack<EssaisUndoSnapshot> essaisUndoSnapshots = new();
        private bool isRestoringAgendaState;

        public event PropertyChangedEventHandler? PropertyChanged;

        public ObservableCollection<Projet> Projets { get; }
        public ICollectionView ProjetsView { get; }
        public ICollectionView ProjetsEssaisView { get; }

        public List<string> Clients { get; }
        public List<string> Demandeurs { get; }
        public List<string> FamillesProduit { get; }
        public List<string> TypesActivite { get; }
        public List<string> Statuts { get; }
        public List<string> NomsEssais { get; }
        public List<SearchFieldOption> ProjetSearchFields { get; }
        public ObservableCollection<AgendaTaskItem> AgendaBacklogTasks { get; }
        public ObservableCollection<AgendaWorkDay> AgendaWeekDays { get; }

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

        public IEnumerable<AgendaTaskItem> AgendaBacklogVisibleTasks =>
            AgendaBacklogTasks
                .Where(task => !AgendaWeekDays.Any(day =>
                    day.PlannedTasks.Any(planned => string.Equals(planned.TaskKey, task.TaskKey, StringComparison.OrdinalIgnoreCase))))
                .ToList();

        public bool CanUndoAgenda => agendaUndoSnapshots.Count > 0;
        public bool CanUndoEssaisBulkAction => essaisUndoSnapshots.Count > 0;

        public int AgendaBacklogVisibleCount => AgendaBacklogVisibleTasks.Count();

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
                essai.Statut = "À faire";
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

        public MainViewModel()
        {
            Clients = new List<string> { "Orange", "Free", "Bouygues", "DTAG", "BT", "N/A" };
            Demandeurs = new List<string> { "RUC", "JEN", "KYJ", "JLC", "JER", "JEL", "JYM", "XAL" };
            FamillesProduit = new List<string> { "Câble", "Cordon", "Autre" };
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

            Projets.CollectionChanged += Projets_CollectionChanged;

            foreach (Projet projet in Projets)
            {
                AttacherProjet(projet);
            }

            RefreshStatistics();
            EnsureSelectedProjetEssais();
            RefreshAgendaTasks();
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
            if (e.PropertyName == nameof(Projet.NomProduit) || e.PropertyName == nameof(Projet.Statut))
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
            RefreshSelectedProjectStatistics();
            RefreshAgendaTasks();
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

        private void RefreshEssaiCollections()
        {
            OnPropertyChanged(nameof(EssaisPreQualificationFiltres));
            OnPropertyChanged(nameof(EssaisQualificationFiltres));
            OnPropertyChanged(nameof(IsEssaiToProcessFilterActive));
            OnPropertyChanged(nameof(IsEssaiInProgressFilterActive));
            OnPropertyChanged(nameof(IsEssaiDoneFilterActive));
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

            if (!isPreQualificationInProgress && !isQualificationInProgress)
            {
                return;
            }

            List<EssaiSuivi> essaisPreQualification = projet.EssaisPreQualification.ToList();
            if (essaisPreQualification.Count == 0)
            {
                return;
            }

            bool allPreQualificationEssaisCompleted = essaisPreQualification.All(essai =>
                !essai.EstConcerne || IsEssaiDoneAndOk(essai));

            if (allPreQualificationEssaisCompleted && isPreQualificationInProgress)
            {
                projet.Statut = "Qualification en cours";
            }
            else if (!allPreQualificationEssaisCompleted && isQualificationInProgress)
            {
                projet.Statut = "Pré-qualification en cours";
            }
        }

        public void MoveAgendaTaskToDay(AgendaTaskItem task, AgendaWorkDay targetDay, double? dropY = null)
        {
            SaveAgendaUndoSnapshot();
            RemoveAgendaTaskFromAllBuckets(task);
            task.ScheduledStartMinutes = GetScheduledStartMinutes(targetDay, task, dropY);
            targetDay.PlannedTasks.Add(task);
            RecalculateAgendaWeek();
            RefreshAgendaCollections();
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

            AgendaWeekTitle = $"Semaine du {monday:dd/MM} au {monday.AddDays(4):dd/MM}";
            return result;
        }

        private void AgendaWorkDay_PropertyChanged(object? sender, PropertyChangedEventArgs e)
        {
            if (isRestoringAgendaState)
            {
                return;
            }

            if (sender is AgendaWorkDay day)
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
            }
        }

        private void RefreshAgendaTasks()
        {
            Dictionary<string, AgendaTaskItem> existingTasks = AgendaBacklogTasks
                .Concat(AgendaWeekDays.SelectMany(day => day.PlannedTasks))
                .ToDictionary(task => task.TaskKey, StringComparer.OrdinalIgnoreCase);

            HashSet<string> validKeys = new(StringComparer.OrdinalIgnoreCase);

            foreach (Projet projet in Projets)
            {
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
            AgendaBacklogTasks.Remove(task);

            foreach (AgendaWorkDay day in AgendaWeekDays)
            {
                day.PlannedTasks.Remove(task);
            }
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

        private void RecalculateAgendaDay(AgendaWorkDay day)
        {
            if (!TryParseTime(day.StartTimeText, out TimeSpan startTime) ||
                !TryParseTime(day.EndTimeText, out TimeSpan endTime) ||
                endTime <= startTime)
            {
                PopulateHourSlots(day, TimeSpan.FromHours(8), TimeSpan.FromHours(16));

                foreach (AgendaTaskItem task in day.PlannedTasks)
                {
                    task.TimeRangeLabel = "Horaires invalides";
                    task.IsOverflow = true;
                    task.TimelineTop = 0;
                    task.BlockHeight = 52;
                    task.TimelineMargin = new Thickness(10, 0, 10, 0);
                }

                return;
            }

            TimeSpan lunchStart = TimeSpan.Zero;
            TimeSpan lunchEnd = TimeSpan.Zero;

            bool hasLunchBreak =
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

            TimeSpan displayStart = RoundDownToHour(startTime);
            TimeSpan displayEnd = RoundUpToHour(endTime);
            if (displayEnd <= displayStart)
            {
                displayEnd = displayStart.Add(TimeSpan.FromHours(1));
            }

            PopulateHourSlots(day, displayStart, displayEnd);

            TimeSpan current = startTime;

            foreach (AgendaTaskItem task in day.PlannedTasks.OrderBy(item => item.ScheduledStartMinutes ?? int.MaxValue))
            {
                TimeSpan taskStart = task.ScheduledStartMinutes.HasValue
                    ? TimeSpan.FromMinutes(task.ScheduledStartMinutes.Value)
                    : current;

                if (taskStart < startTime)
                {
                    taskStart = startTime;
                }

                if (hasLunchBreak && taskStart >= lunchStart && taskStart < lunchEnd)
                {
                    taskStart = lunchEnd;
                }
                TimeSpan taskEnd = ComputeTaskEndTime(taskStart, task.DureeHeures, lunchStart, lunchEnd, hasLunchBreak);

                task.TimeRangeLabel = day.IsWorkingDay
                    ? $"{taskStart:hh\\:mm} - {taskEnd:hh\\:mm}"
                    : "Journée non travaillée";
                task.IsOverflow = taskEnd > endTime;
                task.TimelineTop = Math.Max(0, (taskStart - displayStart).TotalHours * AgendaHourSlotHeight);
                task.BlockHeight = Math.Max(30, ((taskEnd - taskStart).TotalHours * AgendaHourSlotHeight) - 4);
                task.TimelineMargin = new Thickness(10, task.TimelineTop, 10, 0);

                if (!task.ScheduledStartMinutes.HasValue)
                {
                    current = taskEnd;
                }
            }
        }

        private void RecalculateAgendaWeek()
        {
            foreach (AgendaWorkDay day in AgendaWeekDays)
            {
                day.DisplaySegments.Clear();

                if (!TryGetAgendaDayLayout(day, out _, out _, out _, out _, out TimeSpan displayStart, out TimeSpan displayEnd))
                {
                    PopulateHourSlots(day, TimeSpan.FromHours(8), TimeSpan.FromHours(16));
                    continue;
                }

                PopulateHourSlots(day, displayStart, displayEnd);
            }

            for (int dayIndex = 0; dayIndex < AgendaWeekDays.Count; dayIndex++)
            {
                AgendaWorkDay anchorDay = AgendaWeekDays[dayIndex];

                foreach (AgendaTaskItem task in anchorDay.PlannedTasks.OrderBy(item => item.ScheduledStartMinutes ?? int.MaxValue))
                {
                    PlaceTaskAcrossWeek(dayIndex, task);
                }
            }
        }

        private void PlaceTaskAcrossWeek(int startDayIndex, AgendaTaskItem task)
        {
            double remainingHours = task.DureeHeures;
            bool firstSegment = true;
            AgendaTaskSegment? lastSegment = null;

            for (int currentDayIndex = startDayIndex; currentDayIndex < AgendaWeekDays.Count && remainingHours > 0; currentDayIndex++)
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

                TimeSpan segmentStart = firstSegment && task.ScheduledStartMinutes.HasValue
                    ? TimeSpan.FromMinutes(task.ScheduledStartMinutes.Value)
                    : startTime;

                if (segmentStart < startTime)
                {
                    segmentStart = startTime;
                }

                if (segmentStart >= endTime)
                {
                    continue;
                }

                if (segmentStart >= lunchStart && segmentStart < lunchEnd)
                {
                    segmentStart = lunchEnd;
                }

                double availableHours = GetWorkingHoursBetween(segmentStart, endTime, lunchStart, lunchEnd);
                if (availableHours <= 0)
                {
                    continue;
                }

                double segmentHours = Math.Min(remainingHours, availableHours);
                TimeSpan segmentEnd = ComputeTaskEndTime(segmentStart, segmentHours, lunchStart, lunchEnd, lunchEnd > lunchStart);
                double timelineTop = Math.Max(0, (segmentStart - displayStart).TotalHours * AgendaHourSlotHeight);
                double blockHeight = Math.Max(30, ((segmentEnd - segmentStart).TotalHours * AgendaHourSlotHeight) - 4);

                lastSegment = new AgendaTaskSegment
                {
                    SourceTask = task,
                    NomEssai = firstSegment ? task.NomEssai : $"{task.NomEssai} (suite)",
                    NumeroProjet = task.NumeroProjet,
                    NomProduit = task.NomProduit,
                    DureeLabel = $"{segmentHours / AgendaHoursPerDay:0.##} j",
                    TimeRangeLabel = $"{segmentStart:hh\\:mm} - {segmentEnd:hh\\:mm}",
                    IsOverflow = false,
                    BlockHeight = blockHeight,
                    TimelineMargin = new Thickness(10, timelineTop, 10, 0),
                    IsContinuation = !firstSegment
                };

                currentDay.DisplaySegments.Add(lastSegment);

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
            }

            if (remainingHours > 0)
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

        private static double GetWorkingHoursBetween(TimeSpan start, TimeSpan end, TimeSpan lunchStart, TimeSpan lunchEnd)
        {
            if (end <= start)
            {
                return 0;
            }

            double totalHours = (end - start).TotalHours;

            if (lunchEnd > lunchStart)
            {
                TimeSpan overlapStart = start > lunchStart ? start : lunchStart;
                TimeSpan overlapEnd = end < lunchEnd ? end : lunchEnd;

                if (overlapEnd > overlapStart)
                {
                    totalHours -= (overlapEnd - overlapStart).TotalHours;
                }
            }

            return Math.Max(0, totalHours);
        }

        private static bool TryParseTime(string? value, out TimeSpan time)
        {
            return TimeSpan.TryParse(value, CultureInfo.InvariantCulture, out time) ||
                   TimeSpan.TryParse(value, CultureInfo.GetCultureInfo("fr-FR"), out time);
        }

        private static TimeSpan ComputeTaskEndTime(
            TimeSpan taskStart,
            double durationHours,
            TimeSpan lunchStart,
            TimeSpan lunchEnd,
            bool hasLunchBreak)
        {
            TimeSpan current = taskStart;
            double remainingHours = durationHours;

            if (hasLunchBreak && current < lunchStart)
            {
                double hoursBeforeLunch = (lunchStart - current).TotalHours;
                if (remainingHours > hoursBeforeLunch)
                {
                    remainingHours -= hoursBeforeLunch;
                    current = lunchEnd;
                }
                else
                {
                    return current.Add(TimeSpan.FromHours(remainingHours));
                }
            }

            if (hasLunchBreak && current >= lunchStart && current < lunchEnd)
            {
                current = lunchEnd;
            }

            return current.Add(TimeSpan.FromHours(remainingHours));
        }

        private static TimeSpan RoundDownToHour(TimeSpan value)
        {
            return TimeSpan.FromHours(Math.Floor(value.TotalHours));
        }

        private static TimeSpan RoundUpToHour(TimeSpan value)
        {
            return TimeSpan.FromHours(Math.Ceiling(value.TotalHours));
        }

        private int? GetScheduledStartMinutes(AgendaWorkDay day, AgendaTaskItem task, double? dropY)
        {
            if (!dropY.HasValue ||
                !TryParseTime(day.StartTimeText, out TimeSpan startTime) ||
                !TryParseTime(day.EndTimeText, out TimeSpan endTime) ||
                endTime <= startTime)
            {
                return task.ScheduledStartMinutes;
            }

            TimeSpan lunchStart = TimeSpan.Zero;
            TimeSpan lunchEnd = TimeSpan.Zero;

            bool hasLunchBreak =
                TryParseTime(day.LunchStartText, out lunchStart) &&
                TryParseTime(day.LunchEndText, out lunchEnd) &&
                lunchEnd > lunchStart &&
                lunchStart >= startTime &&
                lunchEnd <= endTime;

            TimeSpan displayStart = RoundDownToHour(startTime);
            TimeSpan displayEnd = RoundUpToHour(endTime);
            double totalDisplayHours = Math.Max(1, (displayEnd - displayStart).TotalHours);

            double limitedY = Math.Max(0, Math.Min(dropY.Value, totalDisplayHours * AgendaHourSlotHeight));
            double rawMinutes = displayStart.TotalMinutes + (limitedY / AgendaHourSlotHeight * 60);
            int snappedMinutes = (int)(Math.Round(rawMinutes / 15.0) * 15.0);

            TimeSpan candidate = TimeSpan.FromMinutes(snappedMinutes);

            if (candidate < startTime)
            {
                candidate = startTime;
            }

            if (candidate > endTime)
            {
                candidate = endTime;
            }

            if (hasLunchBreak && candidate >= lunchStart && candidate < lunchEnd)
            {
                candidate = lunchEnd;
            }

            return (int)candidate.TotalMinutes;
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

        private sealed record AgendaTaskPlacementSnapshot(string TaskKey, int? DayIndex, int? ScheduledStartMinutes, int? OrderIndex);

        private sealed record AgendaUndoSnapshot(IReadOnlyList<AgendaTaskPlacementSnapshot> Placements);

        private sealed record EssaiStateSnapshot(int Index, string? Statut, string? ResultatTraitement);

        private sealed record EssaisUndoSnapshot(Projet Projet, IReadOnlyList<EssaiStateSnapshot> States);

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
