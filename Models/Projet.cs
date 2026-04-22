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
            set => SetDateDebutText(value);
        }

        private string? datePrevisionnelle;
        public string? DatePrevisionnelle
        {
            get => datePrevisionnelle;
            set => SetDatePrevisionnelleText(value);
        }

        private string? dateFin;
        public string? DateFin
        {
            get => dateFin;
            set => SetDateFinText(value);
        }

        [JsonIgnore]
        public DateTime? DateDebutValue
        {
            get => ParseDate(dateDebut);
            set => SetDateDebutValue(value);
        }

        [JsonIgnore]
        public DateTime? DatePrevisionnelleValue
        {
            get => ParseDate(datePrevisionnelle);
            set => SetDatePrevisionnelleValue(value);
        }

        [JsonIgnore]
        public DateTime? DateFinValue
        {
            get => ParseDate(dateFin);
            set => SetDateFinValue(value);
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

        private void SetDateDebutText(string? value)
        {
            if (!SetDateText(ref dateDebut, value, nameof(DateDebut), nameof(DateDebutValue)))
            {
                return;
            }

            EnsureDatesNotBeforeStart();
        }

        private void SetDatePrevisionnelleText(string? value)
        {
            string normalizedValue = CoerceDateTextAgainstStart(value);
            SetDateText(ref datePrevisionnelle, normalizedValue, nameof(DatePrevisionnelle), nameof(DatePrevisionnelleValue));
        }

        private void SetDateFinText(string? value)
        {
            string normalizedValue = CoerceDateTextAgainstStart(value);
            SetDateText(ref dateFin, normalizedValue, nameof(DateFin), nameof(DateFinValue));
        }

        private void SetDateDebutValue(DateTime? value)
        {
            if (!SetDateValue(ref dateDebut, value, nameof(DateDebut), nameof(DateDebutValue)))
            {
                return;
            }

            EnsureDatesNotBeforeStart();
        }

        private void SetDatePrevisionnelleValue(DateTime? value)
        {
            SetDateValue(
                ref datePrevisionnelle,
                CoerceDateAgainstStart(value),
                nameof(DatePrevisionnelle),
                nameof(DatePrevisionnelleValue));
        }

        private void SetDateFinValue(DateTime? value)
        {
            SetDateValue(
                ref dateFin,
                CoerceDateAgainstStart(value),
                nameof(DateFin),
                nameof(DateFinValue));
        }

        private bool SetDateText(ref string? backingField, string? value, string textPropertyName, string valuePropertyName)
        {
            string normalizedValue = NormalizeDateText(value);

            if (backingField == normalizedValue)
            {
                return false;
            }

            backingField = normalizedValue;
            OnPropertyChanged(textPropertyName);
            OnPropertyChanged(valuePropertyName);
            return true;
        }

        private bool SetDateValue(ref string? backingField, DateTime? value, string textPropertyName, string valuePropertyName)
        {
            string formattedValue = FormatDate(value);

            if (backingField == formattedValue)
            {
                return false;
            }

            backingField = formattedValue;
            OnPropertyChanged(textPropertyName);
            OnPropertyChanged(valuePropertyName);
            return true;
        }

        private void EnsureDatesNotBeforeStart()
        {
            DateTime? dateDebutValue = ParseDate(dateDebut);
            if (dateDebutValue is null)
            {
                return;
            }

            CoerceStoredDateAgainstStart(ref datePrevisionnelle, nameof(DatePrevisionnelle), nameof(DatePrevisionnelleValue), dateDebutValue.Value);
            CoerceStoredDateAgainstStart(ref dateFin, nameof(DateFin), nameof(DateFinValue), dateDebutValue.Value);
        }

        private void CoerceStoredDateAgainstStart(ref string? backingField, string textPropertyName, string valuePropertyName, DateTime dateDebutValue)
        {
            DateTime? currentValue = ParseDate(backingField);
            if (currentValue is null || currentValue.Value.Date >= dateDebutValue.Date)
            {
                return;
            }

            SetDateValue(ref backingField, dateDebutValue.Date, textPropertyName, valuePropertyName);
        }

        private string CoerceDateTextAgainstStart(string? value)
        {
            string normalizedValue = NormalizeDateText(value);
            DateTime? parsedDate = ParseDate(normalizedValue);

            if (parsedDate is null)
            {
                return normalizedValue;
            }

            return FormatDate(CoerceDateAgainstStart(parsedDate));
        }

        private DateTime? CoerceDateAgainstStart(DateTime? value)
        {
            DateTime? dateDebutValue = ParseDate(dateDebut);
            if (value is null || dateDebutValue is null)
            {
                return value;
            }

            return value.Value.Date < dateDebutValue.Value.Date
                ? dateDebutValue.Value.Date
                : value.Value.Date;
        }

        private static string NormalizeDateText(string? value)
        {
            return string.IsNullOrWhiteSpace(value) ? string.Empty : value.Trim();
        }

        private static string FormatDate(DateTime? value)
        {
            return value?.ToString(DateFormat, DateCulture) ?? string.Empty;
        }
    }
}
