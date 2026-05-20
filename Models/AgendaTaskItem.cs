using System;
using System.ComponentModel;
using System.Windows;

namespace MonTableurApp.Models
{
    public class AgendaTaskItem : INotifyPropertyChanged
    {
        private const int MaxNombrePassages = 5;
        private string? timeRangeLabel;
        private bool isOverflow;
        private double timelineTop;
        private double blockHeight = 52;
        private int? scheduledStartMinutes;
        private bool hasCustomDureeHeures;
        private double dureeMiseEnPlaceHeures;
        private double dureeEssaiHeures;
        private double dureeRepriseHeures;
        private int nombrePassages = 1;
        private bool estArrierePlan;
        private Thickness timelineMargin = new(10, 0, 10, 0);

        public event PropertyChangedEventHandler? PropertyChanged;

        public string TaskKey { get; set; } = string.Empty;
        public string NumeroProjet { get; set; } = string.Empty;
        public string NomProduit { get; set; } = string.Empty;
        public string NomEssai { get; set; } = string.Empty;
        private double dureeJours;

        public double DureeJours
        {
            get => dureeJours;
            set
            {
                if (dureeJours == value)
                {
                    return;
                }

                dureeJours = value;
                OnPropertyChanged(nameof(DureeJours));
                OnPropertyChanged(nameof(DureeLabel));
            }
        }

        public double DureeHeures
        {
            get => DureeMiseEnPlaceHeures +
                   (DureeEssaiHeures * NombrePassages) +
                   (DureeRepriseHeures * Math.Max(0, NombrePassages - 1));
            set
            {
                DureeEssaiHeures = value;
            }
        }

        public double DureeMiseEnPlaceHeures
        {
            get => dureeMiseEnPlaceHeures;
            set
            {
                if (dureeMiseEnPlaceHeures == value)
                {
                    return;
                }

                dureeMiseEnPlaceHeures = value;
                DureeJours = DureeHeures / 7.0;
                NotifyDurationPropertiesChanged();
            }
        }

        public double DureeEssaiHeures
        {
            get => dureeEssaiHeures;
            set
            {
                if (dureeEssaiHeures == value)
                {
                    return;
                }

                dureeEssaiHeures = value;
                DureeJours = DureeHeures / 7.0;
                NotifyDurationPropertiesChanged();
            }
        }

        public double DureeRepriseHeures
        {
            get => dureeRepriseHeures;
            set
            {
                if (dureeRepriseHeures == value)
                {
                    return;
                }

                dureeRepriseHeures = value;
                DureeJours = DureeHeures / 7.0;
                NotifyDurationPropertiesChanged();
            }
        }

        public int NombrePassages
        {
            get => nombrePassages;
            set
            {
                int normalizedValue = Math.Min(MaxNombrePassages, Math.Max(1, value));
                if (nombrePassages == normalizedValue)
                {
                    return;
                }

                nombrePassages = normalizedValue;
                DureeJours = DureeHeures / 7.0;
                NotifyDurationPropertiesChanged();
            }
        }

        public bool EstArrierePlan
        {
            get => estArrierePlan;
            set
            {
                if (estArrierePlan == value)
                {
                    return;
                }

                estArrierePlan = value;
                OnPropertyChanged(nameof(EstArrierePlan));
                OnPropertyChanged(nameof(DureeLabel));
            }
        }

        public string DureeLabel
        {
            get
            {
                string setupLabel = AgendaDurationFormatter.Format(DureeMiseEnPlaceHeures, DureeMiseEnPlaceHeures / 7.0);
                string testLabel = AgendaDurationFormatter.Format(DureeEssaiHeures, DureeEssaiHeures / 7.0);
                string label = EstArrierePlan
                    ? $"Mise {setupLabel} - Essai {testLabel} - fond"
                    : $"Mise {setupLabel} - Essai {testLabel}";

                if (NombrePassages > 1)
                {
                    string repriseLabel = AgendaDurationFormatter.Format(DureeRepriseHeures, DureeRepriseHeures / 7.0);
                    label += $" - x{NombrePassages} - Reprise {repriseLabel}";
                }

                return label;
            }
        }

        public int? ScheduledStartMinutes
        {
            get => scheduledStartMinutes;
            set
            {
                if (scheduledStartMinutes == value)
                {
                    return;
                }

                scheduledStartMinutes = value;
                OnPropertyChanged(nameof(ScheduledStartMinutes));
            }
        }

        public bool HasCustomDureeHeures
        {
            get => hasCustomDureeHeures;
            set
            {
                if (hasCustomDureeHeures == value)
                {
                    return;
                }

                hasCustomDureeHeures = value;
                OnPropertyChanged(nameof(HasCustomDureeHeures));
            }
        }

        public string? TimeRangeLabel
        {
            get => timeRangeLabel;
            set
            {
                if (timeRangeLabel == value)
                {
                    return;
                }

                timeRangeLabel = value;
                OnPropertyChanged(nameof(TimeRangeLabel));
            }
        }

        public bool IsOverflow
        {
            get => isOverflow;
            set
            {
                if (isOverflow == value)
                {
                    return;
                }

                isOverflow = value;
                OnPropertyChanged(nameof(IsOverflow));
            }
        }

        public double TimelineTop
        {
            get => timelineTop;
            set
            {
                if (timelineTop == value)
                {
                    return;
                }

                timelineTop = value;
                OnPropertyChanged(nameof(TimelineTop));
            }
        }

        public double BlockHeight
        {
            get => blockHeight;
            set
            {
                if (blockHeight == value)
                {
                    return;
                }

                blockHeight = value;
                OnPropertyChanged(nameof(BlockHeight));
            }
        }

        public Thickness TimelineMargin
        {
            get => timelineMargin;
            set
            {
                if (timelineMargin == value)
                {
                    return;
                }

                timelineMargin = value;
                OnPropertyChanged(nameof(TimelineMargin));
            }
        }

        private void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        private void NotifyDurationPropertiesChanged()
        {
            OnPropertyChanged(nameof(DureeMiseEnPlaceHeures));
            OnPropertyChanged(nameof(DureeEssaiHeures));
            OnPropertyChanged(nameof(DureeRepriseHeures));
            OnPropertyChanged(nameof(NombrePassages));
            OnPropertyChanged(nameof(DureeHeures));
            OnPropertyChanged(nameof(DureeJours));
            OnPropertyChanged(nameof(DureeLabel));
        }
    }
}
