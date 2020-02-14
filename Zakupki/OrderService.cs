using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Packaging;
using System.Linq;
using System.Net;
using System.Net.Mime;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

using AngleSharp;
using AngleSharp.Dom;
using DocumentFormat.OpenXml.Packaging;
using Microsoft.WindowsAPICodePack.Shell.PropertySystem;

namespace Zakupki
{
    internal class OrderService : IDisposable
    {
        private const string BASE_URL = "https://zakupki.gov.ru/epz";

        private const string SEARCH_ORDER = "/order/extendedsearch/results.html";
        private const string ORDER_DOCUMENTS = "/order/notice/ea44/view/documents.html";
        private const string ORDER_RESULTS = "/order/notice/rpec/search-results.html";

        private const string REG_NUMBER_PARAM = "regNumber";
        private const string ORDER_NUMBER_PARAM = "orderNum";

        private readonly IBrowsingContext _browsingContext;
        private WebClient _webClient;

        private readonly string _searchAddressPattern = $"{BASE_URL}{SEARCH_ORDER}?{{0}}&recordsPerPage=_500&fz44=on&sortBy=PUBLISH_DATE&publishDateFrom=01.01.2018";
        private readonly string _orderResultAddressPattern = $"{BASE_URL}{ORDER_RESULTS}?{ORDER_NUMBER_PARAM}={{0}}";
        private readonly string _orderDocumentsAddressPattern = $"{BASE_URL}{ORDER_DOCUMENTS}?{REG_NUMBER_PARAM}={{0}}";

        public event EventHandler<string> BrowsePage;

        public OrderService()
        {
            IConfiguration config = Configuration.Default.WithDefaultLoader();
            _browsingContext = BrowsingContext.New(config);
        }

        public Order[] SearchOrders(string searchParameter)
        {
            string searchAddress = string.Format(_searchAddressPattern, searchParameter);
            BrowsePage(this, searchAddress);
            IDocument document = _browsingContext.OpenAsync(searchAddress).Result;
            IHtmlCollection<IElement> cells = document.QuerySelectorAll(".search-registry-entry-block");
            Order[] orders = new Order[cells.Length];

            for (int i = 0; i < cells.Length; i++)
            {
                try
                {
                    IElement cell = cells[i];
                    orders[i] = new Order
                    {
                        Number = Regex.Replace(cell.QuerySelector(".registry-entry__header-mid__number a")?.TextContent?.Replace("\n", "").Replace("№", ""), @"\s+", " ").Trim(),
                        Type = Regex.Replace(cell.QuerySelector(".registry-entry__header-top__title")?.TextContent?.Replace("\n", ""), @"\s+", " ").Trim(),
                        State = Regex.Replace(cell.QuerySelector(".registry-entry__header-mid__title")?.TextContent?.Replace("\n", ""), @"\s+", " ").Trim(),
                        Content = Regex.Replace(cell.QuerySelector(".registry-entry__body-value")?.TextContent?.Replace("\n", ""), @"\s+", " ").Trim(),
                        Customer = Regex.Replace(cell.QuerySelector(".registry-entry__body-href a")?.TextContent?.Replace("\n", ""), @"\s+", " ").Trim(),
                        StartPrice = Regex.Replace(cell.QuerySelector(".price-block__value")?.TextContent?.Replace("\n", ""), @"\s+", " ").Trim(),
                    };

                    IHtmlCollection<IElement> datas = cell.QuerySelectorAll(".data-block__value");
                    if (datas != null && datas.Length > 1)
                    {
                        if (DateTime.TryParse(datas[0].TextContent, out DateTime startedAt))
                        {
                            orders[i].StartedAt = startedAt;
                        }
                        if (DateTime.TryParse(datas[1].TextContent, out DateTime updatedAt))
                        {
                            orders[i].UpdatedAt = updatedAt;
                        }
                    }
                }
                catch { }
            }

            return orders;
        }

        public void GetOrderResult(Order order)
        {
            if (order == null)
                return;

            string orderResultAdress = string.Format(_orderResultAddressPattern, order.Number);
            BrowsePage(this, orderResultAdress);
            IDocument document = _browsingContext.OpenAsync(orderResultAdress).Result;

            IHtmlCollection<IElement> tableItems = document.QuerySelectorAll("td.table__row-item");
            if (tableItems != null && tableItems.Length >= 5)
            {
                order.Supplier = tableItems[2].TextContent;
                order.ResultPrice = tableItems[4].TextContent;
            }
        }
        public string DownloadOrderNotice(Order order, string folder)
        {
            if (order == null || string.IsNullOrEmpty(folder))
                return null;

            if (_webClient == null)
                _webClient = new WebClient();

            string orderDocumentsAddress = string.Format(_orderDocumentsAddressPattern, order.Number);
            BrowsePage(this, orderDocumentsAddress);
            IDocument document = _browsingContext.OpenAsync(orderDocumentsAddress).Result;

            string docReference = document.QuerySelector(".notice-documents .attachment .w-wrap-break-word a")?.GetAttribute("href");
            if (docReference == null)
                return null;

            _webClient.Headers.Add("user-agent", "Chrome");
            byte[] data = _webClient.DownloadData(docReference);

            string fileExtension = ".tmp";
            string contentDispositionString = _webClient.ResponseHeaders["Content-Disposition"];
            if (!string.IsNullOrEmpty(contentDispositionString))
            {
                try
                {
                    ContentDisposition contentDisposition = new ContentDisposition(contentDispositionString);
                    string fileName = Uri.UnescapeDataString(contentDisposition.FileName);
                    fileExtension = Path.GetExtension(fileName);
                    order.NoticeName = Path.GetFileNameWithoutExtension(fileName);
                    order.NoticeExtension = fileExtension;
                }
                catch { }
            }
            string fileNameForSave = $"{order.UpdatedAt.ToString("yyyy-MM-dd")}_{order.Number}{fileExtension}";
            string filePathForSave = Path.Combine(folder, fileNameForSave);

            File.WriteAllBytes(filePathForSave, data);

            return filePathForSave;
        }
        public void GetOrderNoticeDetails(Order order, string filePath)
        {
            if (order == null || string.IsNullOrEmpty(filePath))
                return;

            if (string.Equals(order.NoticeExtension, ".doc", StringComparison.InvariantCultureIgnoreCase))
            {
                ProcessDocFile(order, filePath);
            }
            else if (string.Equals(order.NoticeExtension, ".docx", StringComparison.InvariantCultureIgnoreCase))
            {
                ProcessDocxFile(order, filePath);
            }
        }
        private static void ProcessDocxFile(Order order, string filePath)
        {
            using (WordprocessingDocument package = WordprocessingDocument.Open(filePath, true))
            {
                PackageProperties fileProperties = package.PackageProperties;
                order.NoticeCreator = fileProperties.Creator;
                order.NoticeSubject = fileProperties.Subject;
                order.NoticeTitle = fileProperties.Title;
                order.NoticeLastAuthor = fileProperties.LastModifiedBy;
                order.NoticeDatePrinted = fileProperties.LastPrinted.ToString();
                order.NoticeDateCreated = fileProperties.Created.ToString();
                order.NoticeDateSaved = fileProperties.Modified.ToString();
            }
        }
        private static void ProcessDocFile(Order order, string filePath)
        {
            ShellPropertyCollection fileProperties = new ShellPropertyCollection(filePath);

            foreach (var property in fileProperties)
            {
                switch (property.CanonicalName)
                {
                    case "System.Company":
                        string company = property.ValueAsObject?.ToString();
                        order.NoticeCompany = company;
                        break;
                    case "System.Title":
                        string title = property.ValueAsObject?.ToString();
                        order.NoticeSubject = title;
                        break;
                    case "System.Subject":
                        string subject = property.ValueAsObject?.ToString();
                        order.NoticeSubject = subject;
                        break;
                    case "System.Author":
                        if (property.ValueAsObject is string[] authors)
                        {
                            string authorsString = string.Join(", ", authors);
                            order.NoticeCreator = authorsString;
                        }
                        break;
                    case "System.Document.Template":
                        string template = property.ValueAsObject?.ToString();
                        order.NoticeTemplate = template;
                        break;
                    case "System.Document.LastAuthor":
                        string lastAuthor = property.ValueAsObject?.ToString();
                        order.NoticeLastAuthor = lastAuthor;
                        break;
                    case "System.Document.DatePrinted":
                        string datePrinted = property.ValueAsObject?.ToString();
                        order.NoticeDatePrinted = datePrinted;
                        break;
                    case "System.Document.DateCreated":
                        string dateCreated = property.ValueAsObject?.ToString();
                        order.NoticeDateCreated = dateCreated;
                        break;
                    case "System.Document.DateSaved":
                        string dateSaved = property.ValueAsObject?.ToString();
                        order.NoticeDateSaved = dateSaved;
                        break;
                }
            }
        }

        #region IDisposable Support
        private bool isDisposed = false;
        public void Dispose()
        {
            if (!isDisposed)
            {
                _webClient?.Dispose();
                isDisposed = true;
            }
        }
        #endregion
    }
}
