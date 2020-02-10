using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

using AngleSharp;
using AngleSharp.Dom;
using Microsoft.WindowsAPICodePack.Shell.PropertySystem;

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


            string tempFolder = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName());
            Directory.CreateDirectory(tempFolder);
            Process.Start(tempFolder);
            string logFile = Path.Combine(tempFolder, "log.txt");

            IConfiguration config = Configuration.Default.WithDefaultLoader();

            string searchAddress = $"{BASE_URL}{SEARCH_ORDER}?{SEARCH_TEXT_PARAM}=7825660628&recordsPerPage=_500&fz44=on&sortBy=UPDATE_DATE";
            IBrowsingContext context = BrowsingContext.New(config);
            IDocument document = context.OpenAsync(searchAddress).Result;

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

            using (var client = new WebClient())
            {


                string orderDocumentsAddress = $"{BASE_URL}{ORDER_DOCUMENTS}?{ORDER_NUMBER_PARAM}=";
                foreach (var order in orders)
                {
                    Console.WriteLine($"Order number: {order.Number}");
                    File.AppendAllText(logFile, $"Order number: {order.Number}{Environment.NewLine}");

                    Console.WriteLine($"Order content: {order.Content}");
                    File.AppendAllText(logFile, $"Order content: {order.Content}{Environment.NewLine}");

                    Console.WriteLine($"Order start date: {order.StartedAt}");
                    File.AppendAllText(logFile, $"Order start date: {order.StartedAt}{Environment.NewLine}");

                    Console.WriteLine($"Order update date: {order.UpdatedAt}");
                    File.AppendAllText(logFile, $"Order update date: {order.UpdatedAt}{Environment.NewLine}");

                    Console.WriteLine($"Order price: {order.Price}");
                    File.AppendAllText(logFile, $"Order price: {order.Price}{Environment.NewLine}");

                    Console.WriteLine($"Order type: {order.Type}");
                    File.AppendAllText(logFile, $"Order type: {order.Type}{Environment.NewLine}");

                    Console.WriteLine($"Order state: {order.State}");
                    File.AppendAllText(logFile, $"Order state: {order.State}{Environment.NewLine}");

                    Console.WriteLine($"Order customer: {order.Customer}");
                    File.AppendAllText(logFile, $"Order customer: {order.Customer}{Environment.NewLine}");

                    document = context.OpenAsync($"{orderDocumentsAddress}{order.Number}").Result;
                    string docReference = document.QuerySelector(".notice-documents .attachment .w-wrap-break-word a")?.GetAttribute("href");
                    if (docReference != null)
                    {
                        client.Headers.Add("user-agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.2; .NET CLR 1.0.3705;)");
                        var data = client.DownloadData(docReference);

                        string fileExtension = "tmp";
                        if (!string.IsNullOrEmpty(client.ResponseHeaders["Content-Disposition"]))
                        {
                            string fileName = client.ResponseHeaders["Content-Disposition"].Substring(client.ResponseHeaders["Content-Disposition"].IndexOf("filename=") + 9).Replace("\"", "");
                            fileExtension = Path.GetExtension(fileName);
                        }
                        string tempPath = Path.Combine(tempFolder, Path.ChangeExtension(order.Number, fileExtension));

                        File.WriteAllBytes(tempPath, data);


                        foreach (var prop in new ShellPropertyCollection(tempPath))
                        {
                            if (prop.CanonicalName == "System.Author")
                            {

                                string[] authors = prop.ValueAsObject as string[];
                                if (authors != null)
                                {
                                    Console.WriteLine($"Authors: {string.Join("; ", authors)}");
                                    File.AppendAllText(logFile, $"Authors: {string.Join("; ", authors)}{ Environment.NewLine}");
                                }
                            }
                        }
                    }

                    Console.WriteLine();
                    Console.WriteLine();
                    File.AppendAllText(logFile, Environment.NewLine + Environment.NewLine);
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
        public string DocumentCreator { get; set; }
    }

    class MyWebClient : WebClient
    {
        protected override WebRequest GetWebRequest(Uri address)
        {
            HttpWebRequest request = base.GetWebRequest(address) as HttpWebRequest;
            request.AutomaticDecompression = DecompressionMethods.Deflate | DecompressionMethods.GZip;
            return request;
        }
    }
}
