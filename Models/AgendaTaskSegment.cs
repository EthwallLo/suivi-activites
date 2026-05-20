using System.Windows;
using System.ComponentModel;

namespace MonTableurApp.Models
{
    public class AgendaTaskSegment : INotifyPropertyChanged
    {
        private bool isEditingStartTime;
        private bool isEditingEndTime;
        private int columnIndex;
        private int columnCount = 1;

        public event PropertyChangedEventHandler? PropertyChanged;

        public AgendaTaskItem SourceTask { get; set; } = null!;

        public AgendaWorkDay SourceDay { get; set; } = null!;

        public string NomEssai { get; set; } = string.Empty;

        public string NumeroProjet { get; set; } = string.Empty;

        public string NomProduit { get; set; } = string.Empty;

        public string DureeLabel { get; set; } = string.Empty;

        public string PhaseLabel { get; set; } = string.Empty;

        public string TimeRangeLabel { get; set; } = string.Empty;

        public string StartTimeText { get; set; } = string.Empty;

        public string EndTimeText { get; set; } = string.Empty;

        public bool IsOverflow { get; set; }

        public double BlockHeight { get; set; }

        public double TimelineTop { get; set; }

        public Thickness TimelineMargin { get; set; } = new(10, 0, 10, 0);

        public int StartMinutes { get; set; }

        public int EndMinutes { get; set; }

        public int ColumnIndex
        {
            get => columnIndex;
            set
            {
                if (columnIndex == value)
                {
                    return;
                }

                columnIndex = value;
                OnPropertyChanged(nameof(ColumnIndex));
            }
        }

        public int ColumnCount
        {
            get => columnCount;
            set
            {
                int normalizedValue = value <= 0 ? 1 : value;
                if (columnCount == normalizedValue)
                {
                    return;
                }

                columnCount = normalizedValue;
                OnPropertyChanged(nameof(ColumnCount));
                OnPropertyChanged(nameof(IsCompact));
            }
        }

        public bool IsContinuation { get; set; }

        public bool IsSetupSegment { get; set; }

        public bool IsExecutionSegment { get; set; }

        public bool IsBackgroundExecution { get; set; }

        public bool IsRepriseSegment { get; set; }

        public int PassageIndex { get; set; } = 1;

        public int PassageCount { get; set; } = 1;

        public bool CanEditTimeValues { get; set; } = true;

        public bool CanEditStartTime => CanEditTimeValues && !IsContinuation && (IsSetupSegment || IsExecutionSegment);

        public bool CanEditEndTime => CanEditTimeValues && !IsContinuation && (IsSetupSegment || IsExecutionSegment);

        public bool CanEditTimes => CanEditStartTime || CanEditEndTime;

        public bool IsCompact => BlockHeight < 58 || ColumnCount > 1;

        public bool IsTiny => BlockHeight <= 44;

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
