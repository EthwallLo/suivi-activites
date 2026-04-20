using System.ComponentModel;
using System.Collections.Generic;
using System.Linq;
using System.Text.Json.Serialization;

namespace MonTableurApp.Models
{
    public class EssaiSuivi : INotifyPropertyChanged
    {
        private string? nomEssai;
        private string? statut;
        private List<string> statutsDisponibles = new List<string>();

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
                OnPropertyChanged(nameof(Statut));
                OnPropertyChanged(nameof(EstConcerne));
                OnPropertyChanged(nameof(ProgressionPourcentage));
                OnPropertyChanged(nameof(ProgressionTexte));
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
                OnPropertyChanged(nameof(ProgressionPourcentage));
                OnPropertyChanged(nameof(ProgressionTexte));
            }
        }

        [JsonIgnore]
        public bool EstConcerne => !string.Equals(Statut, "Non concerné", System.StringComparison.OrdinalIgnoreCase);

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
                    .Where(statutDisponible => !string.Equals(statutDisponible, "Non concerné", System.StringComparison.OrdinalIgnoreCase))
                    .ToList();

                if (statutsApplicables.Count <= 1)
                {
                    return 0;
                }

                int index = statutsApplicables.FindIndex(statutDisponible =>
                    string.Equals(statutDisponible, Statut, System.StringComparison.OrdinalIgnoreCase));

                if (index < 0)
                {
                    return 0;
                }

                return (int)System.Math.Round(index * 100.0 / (statutsApplicables.Count - 1));
            }
        }

        [JsonIgnore]
        public string ProgressionTexte => EstConcerne ? $"{ProgressionPourcentage} %" : "NC";

        private void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}
