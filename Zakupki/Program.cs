using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

using AngleSharp;
using AngleSharp.Dom;

namespace Zakupki
{
    public class Program
    {
        static void Main(string[] args)
        {
            string BASE_URL = "https://zakupki.gov.ru/epz";

            string SEARCH_ORDER = "/order/extendedsearch/results.html";
            string SEARCH_TEXT_PARAM = "searchString";

            string ORDER_DOCUMENTS = "/order/notice/ea44/view/documents.html";
            string ORDER_NUMBER_PARAM = "regNumber";

            IConfiguration config = Configuration.Default.WithDefaultLoader();
            string address = $"{BASE_URL}{SEARCH_ORDER}?{SEARCH_TEXT_PARAM}=лиговка-ямская&recordsPerPage=_500";
            IBrowsingContext context = BrowsingContext.New(config);
            IDocument document = context.OpenAsync(address).Result;

            IHtmlCollection<IElement> cells = document.QuerySelectorAll(".search-registry-entry-block");
            Order[] orders = new Order[cells.Length];
            for (int i = 0; i < cells.Length; i++)
            {
                IElement cell = cells[i];
                orders[i] = new Order
                {
                    Number = Regex.Replace(cell.QuerySelector(".registry-entry__header-mid__number a")?.TextContent?.Replace("\n", "").Replace("№", ""), @"\s+", " ").Trim(),
                    Type = Regex.Replace(cell.QuerySelector(".registry-entry__header-top__title")?.TextContent?.Replace("\n", ""), @"\s+", " ").Trim(),
                    State = Regex.Replace(cell.QuerySelector(".registry-entry__header-mid__title")?.TextContent?.Replace("\n", ""), @"\s+", " ").Trim(),
                    Content = Regex.Replace(cell.QuerySelector(".registry-entry__body-value")?.TextContent?.Replace("\n", ""), @"\s+", " ").Trim(),
                    Customer = Regex.Replace(cell.QuerySelector(".registry-entry__body-href a")?.TextContent?.Replace("\n", ""), @"\s+", " ").Trim(),
                    Price = Regex.Replace(cell.QuerySelector(".price-block__value")?.TextContent?.Replace("\n", ""), @"\s+", " ").Trim(),
                };

                IHtmlCollection<IElement> datas = cell.QuerySelectorAll(".data-block__value");
                if (datas != null && datas.Length > 1)
                {
                    orders[i].StartedAt = datas[0].TextContent;
                    orders[i].UpdatedAt = datas[1].TextContent;
                }
            }
        }
    }
    public class Order
    {
        public string Number { get; set; }
        public string Type { get; set; }
        public string State { get; set; }
        public string Content { get; set; }
        public string Customer { get; set; }
        public string Price { get; set; }
        public string StartedAt { get; set; }
        public string UpdatedAt { get; set; }
    }
}
