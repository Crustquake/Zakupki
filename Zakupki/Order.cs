using System;

namespace Zakupki
{
    internal class Order
    {
        public string Number { get; set; }
        public DateTime StartedAt { get; set; }
        public DateTime UpdatedAt { get; set; }
        public string State { get; set; }
        public string Content { get; set; }
        public string Customer { get; set; }
        public string StartPrice { get; set; }
        public string ResultPrice { get; set; }
        public string Supplier { get; set; }
        public string Type { get; set; }
        public string NoticeName { get; set; }
        public string NoticeExtension { get; set; }
        public string NoticeCompany { get; set; }
        public string NoticeCreator { get; set; }
        public string NoticeLastAuthor { get; set; }
        public string NoticeDateCreated { get; set; }
        public string NoticeDateSaved { get; set; }
        public string NoticeDatePrinted { get; set; }
        public string NoticeTemplate { get; set; }
        public string NoticeSubject { get; set; }
        public string NoticeTitle { get; set; }
    }
}
