using System.Windows;
using System.ComponentModel;

namespace MonTableurApp.Models
{
    public class AgendaTaskSegment : INotifyPropertyChanged
    {
        private bool isEditingStartTime;
        private bool isEditingEndTime;

        public event PropertyChangedEventHandler? PropertyChanged;

        public AgendaTaskItem SourceTask { get; set; } = null!;

        public AgendaWorkDay SourceDay { get; set; } = null!;

        public string NomEssai { get; set; } = string.Empty;

        public string NumeroProjet { get; set; } = string.Empty;

        public string NomProduit { get; set; } = string.Empty;

        public string DureeLabel { get; set; } = string.Empty;

        public string TimeRangeLabel { get; set; } = string.Empty;

        public string StartTimeText { get; set; } = string.Empty;

        public string EndTimeText { get; set; } = string.Empty;

        public bool IsOverflow { get; set; }

        public double BlockHeight { get; set; }

        public Thickness TimelineMargin { get; set; } = new(10, 0, 10, 0);

        public bool IsContinuation { get; set; }

        public bool CanEditTimes => !IsContinuation;

        public bool IsEditingStartTime
        {
            get => isEditingStartTime;
            set
            {
                if (isEditingStartTime == value)
                {
                    return;
                }

                isEditingStartTime = value;
                OnPropertyChanged(nameof(IsEditingStartTime));
            }
        }

        public bool IsEditingEndTime
        {
            get => isEditingEndTime;
            set
            {
                if (isEditingEndTime == value)
                {
                    return;
                }

                isEditingEndTime = value;
                OnPropertyChanged(nameof(IsEditingEndTime));
            }
        }

        private void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}
