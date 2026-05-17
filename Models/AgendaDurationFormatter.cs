using System;

namespace MonTableurApp.Models
{
    public static class AgendaDurationFormatter
    {
        public static string Format(double durationHours, double durationDays)
        {
            if (durationDays >= 1)
            {
                return $"{durationDays:0.##} j";
            }

            int totalMinutes = Math.Max(0, (int)Math.Round(durationHours * 60, MidpointRounding.AwayFromZero));
            int hours = totalMinutes / 60;
            int minutes = totalMinutes % 60;

            if (hours > 0 && minutes > 0)
            {
                return $"{hours} h {minutes} min";
            }

            if (hours > 0)
            {
                return $"{hours} h";
            }

            return $"{minutes} min";
        }
    }
}
