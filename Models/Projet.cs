using System.Collections.ObjectModel;
using System.Collections.Generic;
using System.ComponentModel;
using System.Globalization;
using System.Linq;
using System.Text.Json.Serialization;

namespace MonTableurApp.Models
{
    public class Projet : INotifyPropertyChanged
    {
        private static readonly CultureInfo DateCulture = CultureInfo.GetCultureInfo("fr-FR");
        private const string DateFormat = "dd/MM/yyyy";

        private static readonly string[] EssaisPreQualificationNames =
        {
            "Traction 100m",
            "Cyclage thermique",
            "Statique Bending",
            "Vieillissement"
        };

        public event PropertyChangedEventHandler? PropertyChanged;

        private void OnPropertyChanged(string name)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
        }

        private string? numeroProjet;
        public string? NumeroProjet
        {
            get => numeroProjet;
            set { numeroProjet = value; OnPropertyChanged(nameof(NumeroProjet)); }
        }

        private string? nomProduit;
        public string? NomProduit
        {
            get => nomProduit;
            set { nomProduit = value; OnPropertyChanged(nameof(NomProduit)); }
        }

        private string? familleProduit;
        public string? FamilleProduit
        {
            get => familleProduit;
            set { familleProduit = value; OnPropertyChanged(nameof(FamilleProduit)); }
        }

        private string? client;
        public string? Client
        {
            get => client;
            set { client = value; OnPropertyChanged(nameof(Client)); }
        }

        private string? demandeur;
        public string? Demandeur
        {
            get => demandeur;
            set { demandeur = value; OnPropertyChanged(nameof(Demandeur)); }
        }

        private string? typeActivite;
        public string? TypeActivite
        {
            get => typeActivite;
            set { typeActivite = value; OnPropertyChanged(nameof(TypeActivite)); }
        }

        private string? dossierRacine;
        public string? DossierRacine
        {
            get => dossierRacine;
            set { dossierRacine = value; OnPropertyChanged(nameof(DossierRacine)); }
        }

        private string? statut;
        public string? Statut
        {
            get => statut;
            set { statut = value; OnPropertyChanged(nameof(Statut)); }
        }

        private string? dateDebut;
        public string? DateDebut
        {
            get => dateDebut;
            set => SetDateText(ref dateDebut, value, nameof(DateDebut), nameof(DateDebutValue));
        }

        private string? datePrevisionnelle;
        public string? DatePrevisionnelle
        {
            get => datePrevisionnelle;
            set => SetDateText(ref datePrevisionnelle, value, nameof(DatePrevisionnelle), nameof(DatePrevisionnelleValue));
        }

        private string? dateFin;
        public string? DateFin
        {
            get => dateFin;
            set => SetDateText(ref dateFin, value, nameof(DateFin), nameof(DateFinValue));
        }

        [JsonIgnore]
        public DateTime? DateDebutValue
        {
            get => ParseDate(dateDebut);
            set => SetDateValue(ref dateDebut, value, nameof(DateDebut), nameof(DateDebutValue));
        }

        [JsonIgnore]
        public DateTime? DatePrevisionnelleValue
        {
            get => ParseDate(datePrevisionnelle);
            set => SetDateValue(ref datePrevisionnelle, value, nameof(DatePrevisionnelle), nameof(DatePrevisionnelleValue));
        }

        [JsonIgnore]
        public DateTime? DateFinValue
        {
            get => ParseDate(dateFin);
            set => SetDateValue(ref dateFin, value, nameof(DateFin), nameof(DateFinValue));
        }

        private string? commentaires;
        public string? Commentaires
        {
            get => commentaires;
            set { commentaires = value; OnPropertyChanged(nameof(Commentaires)); }
        }

        private ObservableCollection<EssaiSuivi> essais = new ObservableCollection<EssaiSuivi>();
        public ObservableCollection<EssaiSuivi> Essais
        {
            get => essais;
            set
            {
                essais = value ?? new ObservableCollection<EssaiSuivi>();
                OnPropertyChanged(nameof(Essais));
                OnPropertyChanged(nameof(EssaisPreQualification));
                OnPropertyChanged(nameof(EssaisQualification));
            }
        }

        [JsonIgnore]
        public IEnumerable<EssaiSuivi> EssaisPreQualification =>
            Essais.Where(essai => EssaisPreQualificationNames.Contains(essai.NomEssai ?? string.Empty));

        [JsonIgnore]
        public IEnumerable<EssaiSuivi> EssaisQualification =>
            Essais.Where(essai => !EssaisPreQualificationNames.Contains(essai.NomEssai ?? string.Empty));

        private static DateTime? ParseDate(string? value)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                return null;
            }

            if (DateTime.TryParseExact(value.Trim(), DateFormat, DateCulture, DateTimeStyles.None, out DateTime parsedDate))
            {
                return parsedDate;
            }

            return DateTime.TryParse(value, DateCulture, DateTimeStyles.None, out parsedDate)
                ? parsedDate.Date
                : null;
        }

        private void SetDateText(ref string? backingField, string? value, string textPropertyName, string valuePropertyName)
        {
            string normalizedValue = string.IsNullOrWhiteSpace(value) ? string.Empty : value.Trim();

            if (backingField == normalizedValue)
            {
                return;
            }

            backingField = normalizedValue;
            OnPropertyChanged(textPropertyName);
            OnPropertyChanged(valuePropertyName);
        }

        private void SetDateValue(ref string? backingField, DateTime? value, string textPropertyName, string valuePropertyName)
        {
            string formattedValue = value?.ToString(DateFormat, DateCulture) ?? string.Empty;

            if (backingField == formattedValue)
            {
                return;
            }

            backingField = formattedValue;
            OnPropertyChanged(textPropertyName);
            OnPropertyChanged(valuePropertyName);
        }
    }
}
