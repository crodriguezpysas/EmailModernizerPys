using System;

namespace EmailModernizerPys.Models
{
    public class EmailSummary
    {
        public string From { get; set; }
        public DateTime DateTimeReceived { get; set; }
        public string To { get; set; }
        public string Subject { get; set; }
        public string Attachments { get; set; }
    }
}