using System;

namespace SecurityPlus.Models
{
    internal class Duty
    {
        public int Id { get; set; }
        public int VisibleId { get; set; }
        public string Fullname { get; set; }
        public string VisibleFullname { get; set; }
        public bool IsCar { get; set; }

        public string VisibleIsCar => IsCar ? "+" : "-";
        public DateTime DateStart { get; set; } = DateTime.Now;
        public DateTime VisibleDateStart { get; set; } = DateTime.Now;
        public DateTime TimeStart { get; set; }
        public DateTime VisibleTimeStart { get; set; }
        public DateTime DateEnd { get; set; } = DateTime.Now;
        public DateTime VisibleDateEnd { get; set; } = DateTime.Now;
        public DateTime TimeEnd { get; set; }
        public DateTime VisibleTimeEnd { get; set; }
        public int Time { get; set; }
        public string TimeString { get; set; }
        public decimal Sum { get; set; }
        public bool IsPrint { get; set; }
    }
}
