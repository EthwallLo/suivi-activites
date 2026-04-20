using System.Collections.ObjectModel;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text.Json.Serialization;

namespace MonTableurApp.Models
{
    public class Projet : INotifyPropertyChanged
    {
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
            set { dateDebut = value; OnPropertyChanged(nameof(DateDebut)); }
        }

        private string? datePrevisionnelle;
        public string? DatePrevisionnelle
        {
            get => datePrevisionnelle;
            set { datePrevisionnelle = value; OnPropertyChanged(nameof(DatePrevisionnelle)); }
        }

        private string? dateFin;
        public string? DateFin
        {
            get => dateFin;
            set { dateFin = value; OnPropertyChanged(nameof(DateFin)); }
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
    }
}
