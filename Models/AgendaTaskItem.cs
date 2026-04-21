using System.ComponentModel;
using System.Windows;

namespace MonTableurApp.Models
{
    public class AgendaTaskItem : INotifyPropertyChanged
    {
        private string? timeRangeLabel;
        private bool isOverflow;
        private double timelineTop;
        private double blockHeight = 52;
        private int? scheduledStartMinutes;
        private Thickness timelineMargin = new(10, 0, 10, 0);

        public event PropertyChangedEventHandler? PropertyChanged;

        public string TaskKey { get; set; } = string.Empty;
        public string NumeroProjet { get; set; } = string.Empty;
        public string NomProduit { get; set; } = string.Empty;
        public string NomEssai { get; set; } = string.Empty;
        public double DureeJours { get; set; }
        public double DureeHeures { get; set; }
        public string DureeLabel => $"{DureeJours:0.##} j";

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
    }
}
