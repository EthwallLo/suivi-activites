using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
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
        private List<string> statutsDisponibles = new();

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

        private static string NormalizeValue(string? value)
        {
            return (value ?? string.Empty)
                .Trim()
                .ToLowerInvariant()
                .Replace("Ã©", "e")
                .Replace("Ã¨", "e")
                .Replace("Ãª", "e")
                .Replace("é", "e")
                .Replace("è", "e")
                .Replace("ê", "e")
                .Replace("à", "a");
        }
    }
}
