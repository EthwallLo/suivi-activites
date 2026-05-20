using System.Windows;

namespace MonTableurApp.Models
{
    public class AgendaTimeBlock
    {
        public string Label { get; set; } = string.Empty;

        public bool ShowLabel => !string.IsNullOrWhiteSpace(Label) && BlockHeight >= 20;

        public double BlockHeight { get; set; }

        public Thickness TimelineMargin { get; set; } = new(0, 0, 0, 0);

        public CornerRadius CornerRadius { get; set; } = new(0);
    }
}
