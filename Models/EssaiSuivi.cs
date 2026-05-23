using System.Collections.Generic;
using System.ComponentModel;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Text.Json.Serialization;
using System.Windows;
using System.Windows.Media;

namespace MonTableurApp.Models
{
    public class EssaiSuivi : INotifyPropertyChanged
    {
        private string? nomEssai;
        private string? statut;
        private string? resultatTraitement;
        private string? commentaire;
        private string? categorie;
        private string? referenceProduitUtilisee;
        private int nombrePassages;
        private bool hasCustomPlanning;
        private double customDureeMiseEnPlaceHeures;
        private double customDureeEssaiHeures;
        private double customDureeRepriseHeures;
        private bool customEstArrierePlan;
        private List<string> statutsDisponibles = new();
        private const int MaxNombrePassages = 5;
        private static readonly IReadOnlyList<int> PassageCounts = Enumerable.Range(1, MaxNombrePassages).ToList();

        public event PropertyChangedEventHandler? PropertyChanged;

        public string? NomEssai
        {
            get => nomEssai;
            set
            {
                if (nomEssai == value)
                {
                    return;
                }

                nomEssai = value;
                OnPropertyChanged(nameof(NomEssai));
                OnPropertyChanged(nameof(AfficheNombrePassages));
            }
        }

        public string? Statut
        {
            get => statut;
            set
            {
                if (statut == value)
                {
                    return;
                }

                statut = value;
                if (!EstFait)
                {
                    ResultatTraitement = null;
                }

                OnPropertyChanged(nameof(Statut));
                OnPropertyChanged(nameof(EstConcerne));
                OnPropertyChanged(nameof(EstFait));
                OnPropertyChanged(nameof(AfficheIndicateurEtat));
                OnPropertyChanged(nameof(CouleurIndicateurEtat));
                OnPropertyChanged(nameof(ProgressionPourcentage));
                OnPropertyChanged(nameof(ProgressionTexte));
            }
        }

        [JsonIgnore]
        public string? Categorie
        {
            get => categorie;
            set
            {
                if (categorie == value)
                {
                    return;
                }

                categorie = value;
                OnPropertyChanged(nameof(Categorie));
            }
        }

        public string? ResultatTraitement
        {
            get => resultatTraitement;
            set
            {
                if (resultatTraitement == value)
                {
                    return;
                }

                resultatTraitement = value;
                OnPropertyChanged(nameof(ResultatTraitement));
                OnPropertyChanged(nameof(AfficheIndicateurEtat));
                OnPropertyChanged(nameof(CouleurIndicateurEtat));
            }
        }

        public string? Commentaire
        {
            get => commentaire;
            set
            {
                if (commentaire == value)
                {
                    return;
                }

                commentaire = value;
                OnPropertyChanged(nameof(Commentaire));
            }
        }

        public string? ReferenceProduitUtilisee
        {
            get => referenceProduitUtilisee;
            set
            {
                if (referenceProduitUtilisee == value)
                {
                    return;
                }

                referenceProduitUtilisee = value;
                OnPropertyChanged(nameof(ReferenceProduitUtilisee));
                OnPropertyChanged(nameof(ReferenceProduitUtiliseeLabel));
            }
        }

        [JsonIgnore]
        public string ReferenceProduitUtiliseeLabel => string.IsNullOrWhiteSpace(ReferenceProduitUtilisee)
            ? "Réf. à préciser"
            : ReferenceProduitUtilisee;

        public int NombrePassages
        {
            get => System.Math.Min(MaxNombrePassages, nombrePassages <= 0 ? 1 : nombrePassages);
            set
            {
                int normalizedValue = System.Math.Min(MaxNombrePassages, System.Math.Max(1, value));
                if (nombrePassages == normalizedValue)
                {
                    return;
                }

                nombrePassages = normalizedValue;
                OnPropertyChanged(nameof(NombrePassages));
                OnPropertyChanged(nameof(AfficheNombrePassages));
            }
        }

        public bool HasCustomPlanning
        {
            get => hasCustomPlanning;
            set
            {
                if (hasCustomPlanning == value)
                {
                    return;
                }

                hasCustomPlanning = value;
                OnPropertyChanged(nameof(HasCustomPlanning));
                OnPropertyChanged(nameof(PlanningConfigurationLabel));
            }
        }

        public double CustomDureeMiseEnPlaceHeures
        {
            get => customDureeMiseEnPlaceHeures;
            set
            {
                if (System.Math.Abs(customDureeMiseEnPlaceHeures - value) < 0.001)
                {
                    return;
                }

                customDureeMiseEnPlaceHeures = value;
                OnPropertyChanged(nameof(CustomDureeMiseEnPlaceHeures));
            }
        }

        public double CustomDureeEssaiHeures
        {
            get => customDureeEssaiHeures;
            set
            {
                if (System.Math.Abs(customDureeEssaiHeures - value) < 0.001)
                {
                    return;
                }

                customDureeEssaiHeures = value;
                OnPropertyChanged(nameof(CustomDureeEssaiHeures));
            }
        }

        public double CustomDureeRepriseHeures
        {
            get => customDureeRepriseHeures;
            set
            {
                if (System.Math.Abs(customDureeRepriseHeures - value) < 0.001)
                {
                    return;
                }

                customDureeRepriseHeures = value;
                OnPropertyChanged(nameof(CustomDureeRepriseHeures));
            }
        }

        public bool CustomEstArrierePlan
        {
            get => customEstArrierePlan;
            set
            {
                if (customEstArrierePlan == value)
                {
                    return;
                }

                customEstArrierePlan = value;
                OnPropertyChanged(nameof(CustomEstArrierePlan));
            }
        }

        [JsonIgnore]
        public bool HasNombrePassagesDefini => nombrePassages > 0;

        [JsonIgnore]
        public bool AfficheNombrePassages => IsRepeatedEssaiName(NomEssai) || NombrePassages > 1;

        [JsonIgnore]
        public IReadOnlyList<int> NombrePassagesDisponibles => PassageCounts;

        [JsonIgnore]
        public string PlanningConfigurationLabel => HasCustomPlanning ? "Perso" : "Défaut";

        [JsonIgnore]
        public List<string> StatutsDisponibles
        {
            get => statutsDisponibles;
            set
            {
                statutsDisponibles = value ?? new List<string>();
                OnPropertyChanged(nameof(StatutsDisponibles));
                OnPropertyChanged(nameof(AfficheIndicateurEtat));
                OnPropertyChanged(nameof(CouleurIndicateurEtat));
                OnPropertyChanged(nameof(ProgressionPourcentage));
                OnPropertyChanged(nameof(ProgressionTexte));
            }
        }

        [JsonIgnore]
        public bool EstConcerne => NormalizeValue(Statut) != "non concerne";

        [JsonIgnore]
        public bool EstFait => NormalizeValue(Statut) == "fait";

        [JsonIgnore]
        public List<string> ResultatsTraitementDisponibles { get; } = new() { "OK", "NOK" };

        [JsonIgnore]
        public Visibility AfficheIndicateurEtat =>
            HasIndicatorState ? Visibility.Visible : Visibility.Collapsed;

        [JsonIgnore]
        public SolidColorBrush CouleurIndicateurEtat
        {
            get
            {
                string resultat = NormalizeValue(ResultatTraitement);
                if (resultat == "ok")
                {
                    return CreateFrozenBrush("#68C97D");
                }

                if (resultat == "nok")
                {
                    return CreateFrozenBrush("#F27D8E");
                }

                return CreateFrozenBrush("#F1C95C");
            }
        }

        [JsonIgnore]
        public int ProgressionPourcentage
        {
            get
            {
                if (!EstConcerne)
                {
                    return 0;
                }

                List<string> statutsApplicables = StatutsDisponibles
                    .Where(statutDisponible => NormalizeValue(statutDisponible) != "non concerne")
                    .ToList();

                if (statutsApplicables.Count <= 1)
                {
                    return 0;
                }

                int index = statutsApplicables.FindIndex(statutDisponible =>
                    NormalizeValue(statutDisponible) == NormalizeValue(Statut));

                if (index < 0)
                {
                    return 0;
                }

                return (int)System.Math.Round(index * 100.0 / (statutsApplicables.Count - 1));
            }
        }

        [JsonIgnore]
        public string ProgressionTexte => EstConcerne ? $"{ProgressionPourcentage} %" : "NC";

        private bool HasIndicatorState
        {
            get
            {
                if (!EstConcerne)
                {
                    return false;
                }

                string statutNormalise = NormalizeValue(Statut);
                return statutNormalise != string.Empty && statutNormalise != "a faire";
            }
        }

        private void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        private static SolidColorBrush CreateFrozenBrush(string color)
        {
            var brush = (SolidColorBrush)new BrushConverter().ConvertFromString(color)!;
            brush.Freeze();
            return brush;
        }

        private static bool IsRepeatedEssaiName(string? nomEssai)
        {
            return NormalizeValue(nomEssai) is
                "crush" or
                "cut through" or
                "traction pince" or
                "traction spirale" or
                "traction spiralin";
        }

        private static string NormalizeValue(string? value)
        {
            string repairedValue = RepairMojibakeIfNeeded(value);
            string normalized = repairedValue
                .Trim()
                .ToLowerInvariant()
                .Normalize(NormalizationForm.FormD);

            var builder = new StringBuilder(normalized.Length);

            foreach (char current in normalized)
            {
                if (CharUnicodeInfo.GetUnicodeCategory(current) != UnicodeCategory.NonSpacingMark)
                {
                    builder.Append(current);
                }
            }

            return builder
                .ToString()
                .Normalize(NormalizationForm.FormC);
        }

        private static string RepairMojibakeIfNeeded(string? value)
        {
            string input = value ?? string.Empty;
            if (!input.Contains('\u00C3') &&
                !input.Contains('\u00C2') &&
                !input.Contains('\u00E2') &&
                !input.Contains('\uFFFD'))
            {
                return input;
            }

            try
            {
                if (input.Contains('\u00C3') || input.Contains('\u00C2') || input.Contains('\u00E2'))
                {
                    input = Encoding.UTF8.GetString(Encoding.Latin1.GetBytes(input));
                }
            }
            catch (ArgumentException)
            {
            }

            return input
                .Replace("C\uFFFDbles", "Câbles")
                .Replace("C\uFFFDble", "Câble")
                .Replace("activit\uFFFDs", "activités")
                .Replace("Activit\uFFFDs", "Activités")
                .Replace("P\uFFFDtroth\uFFFDne", "Pétrothène");
        }
    }
}

