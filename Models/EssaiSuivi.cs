using System.ComponentModel;
using System.Text.Json.Serialization;
using System.Collections.Generic;

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
            }
        }

        private void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}
