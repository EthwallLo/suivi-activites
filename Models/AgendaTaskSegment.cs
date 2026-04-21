using System.Windows;

namespace MonTableurApp.Models
{
    public class AgendaTaskSegment
    {
        public AgendaTaskItem SourceTask { get; set; } = null!;

        public string NomEssai { get; set; } = string.Empty;

        public string NumeroProjet { get; set; } = string.Empty;

        public string NomProduit { get; set; } = string.Empty;

        public string DureeLabel { get; set; } = string.Empty;

        public string TimeRangeLabel { get; set; } = string.Empty;

        public bool IsOverflow { get; set; }

        public double BlockHeight { get; set; }

        public Thickness TimelineMargin { get; set; } = new(10, 0, 10, 0);

        public bool IsContinuation { get; set; }
    }
}
