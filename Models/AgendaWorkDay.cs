using System;
using System.Collections.ObjectModel;
using System.ComponentModel;

namespace MonTableurApp.Models
{
    public class AgendaWorkDay : INotifyPropertyChanged
    {
        private bool isWorkingDay = true;
        private string startTimeText = "08:00";
        private string lunchStartText = "13:00";
        private string lunchEndText = "13:30";
        private string endTimeText = "16:00";
        private double timelineHeight = 576;
        private string endHourLabel = "16:00";

        public event PropertyChangedEventHandler? PropertyChanged;

        public AgendaWorkDay(string dayLabel, DateTime date)
        {
            DayLabel = dayLabel;
            Date = date;
        }

        public string DayLabel { get; }

        public DateTime Date { get; }

        public string DateLabel => Date.ToString("dd/MM");

        public ObservableCollection<AgendaTaskItem> PlannedTasks { get; } = new();

        public ObservableCollection<AgendaTaskSegment> DisplaySegments { get; } = new();

        public ObservableCollection<string> HourSlots { get; } = new();

        public bool IsWorkingDay
        {
            get => isWorkingDay;
            set
            {
                if (isWorkingDay == value)
                {
                    return;
                }

                isWorkingDay = value;
                OnPropertyChanged(nameof(IsWorkingDay));
            }
        }

        public string StartTimeText
        {
            get => startTimeText;
            set
            {
                if (startTimeText == value)
                {
                    return;
                }

                startTimeText = value;
                OnPropertyChanged(nameof(StartTimeText));
            }
        }

        public string LunchStartText
        {
            get => lunchStartText;
            set
            {
                if (lunchStartText == value)
                {
                    return;
                }

                lunchStartText = value;
                OnPropertyChanged(nameof(LunchStartText));
            }
        }

        public string LunchEndText
        {
            get => lunchEndText;
            set
            {
                if (lunchEndText == value)
                {
                    return;
                }

                lunchEndText = value;
                OnPropertyChanged(nameof(LunchEndText));
            }
        }

        public string EndTimeText
        {
            get => endTimeText;
            set
            {
                if (endTimeText == value)
                {
                    return;
                }

                endTimeText = value;
                OnPropertyChanged(nameof(EndTimeText));
            }
        }

        public double TimelineHeight
        {
            get => timelineHeight;
            set
            {
                if (timelineHeight == value)
                {
                    return;
                }

                timelineHeight = value;
                OnPropertyChanged(nameof(TimelineHeight));
            }
        }

        public string EndHourLabel
        {
            get => endHourLabel;
            set
            {
                if (endHourLabel == value)
                {
                    return;
                }

                endHourLabel = value;
                OnPropertyChanged(nameof(EndHourLabel));
            }
        }

        private void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}
