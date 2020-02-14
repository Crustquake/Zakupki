using System;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;

using OfficeOpenXml;

namespace Zakupki
{
    internal class OrdersProcessing : IDisposable
    {
        private ExcelPackage _excelPackage;
        public void Run()
        {
            Console.WriteLine("Введите параметр поиска, например ИНН, в виде:");
            Console.WriteLine("для простого поиска \"searchString=7842006310\"");
            Console.WriteLine("для поиска по участнику закупки \"participantName=7842148635\"");

            string searchParameter;
            bool isParameterValid = false;
            do
            {
                searchParameter = Console.ReadLine();

                isParameterValid = !string.IsNullOrEmpty(searchParameter) && searchParameter.Contains("=");
                if (!isParameterValid)
                {
                    Console.WriteLine("Параметр поиска некорректный, попробуйте ещё раз");
                }
            }
            while (!isParameterValid);

            Console.WriteLine();

            using (OrderService service = new OrderService())
            {
                service.BrowsePage += (source, adress) => Console.WriteLine($"Открываем страницу \"{adress}\"");

                Order[] orders = service.SearchOrders(searchParameter);

                Console.WriteLine($"{orders.Length} orders was found");
                Console.WriteLine();

                if (orders.Length > 0)
                {
                    string tempFolder = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName());

                    Directory.CreateDirectory(tempFolder);
                    Process.Start("explorer.exe", tempFolder);

                    string resultFilePath = Path.Combine(tempFolder, "result.xlsx");

                    _excelPackage = CreateExcelFile(resultFilePath);

                    //using (CsvWriter csvWriter = new CsvWriter(streamWriter, CultureInfo.InvariantCulture))
                    //{
                    //    //csvWriter.Configuration.Delimiter = ";";
                    //    //csvWriter.WriteHeader(typeof(Order));
                    //    //csvWriter.NextRecord();

                    //    foreach (Order order in orders)
                    //    {
                    //        if (order != null)
                    //        {
                    //            service.GetOrderResult(order);
                    //            var orderNoticeFile = service.DownloadOrderNotice(order, tempFolder);
                    //            if (orderNoticeFile != null)
                    //            {
                    //                service.GetOrderNoticeDetails(order, orderNoticeFile);
                    //            }

                    //            WriteOrderToConsole(order);

                    //            //csvWriter.WriteRecord(order);
                    //            //csvWriter.NextRecord();
                    //        }
                    //    }
                    //}
                    _excelPackage.Save();
                }
            }
            Console.WriteLine("Приложение закончило работу, нажмите любую кнопку чтобы выйти...");
            Console.ReadKey();
        }

        private void WriteOrderToConsole(Order order)
        {
            Console.WriteLine();
            Console.WriteLine($"Номер закупки: {order.Number}");
            Console.WriteLine($"Размещено: {order.StartedAt.ToShortDateString()}");
            Console.WriteLine($"Обновлено: {order.UpdatedAt.ToShortDateString()}");
            Console.WriteLine($"Статус: {order.State}");
            Console.WriteLine($"Объект закупки: {order.Content}");
            Console.WriteLine($"Заказчик: {order.Customer}");
            Console.WriteLine($"Начальная цена: {order.StartPrice}");

            if (order.ResultPrice != null)
                Console.WriteLine($"Цена контракта: {order.ResultPrice}");

            if (order.Supplier != null)
                Console.WriteLine($"Поставщик: {order.Supplier}");

            Console.WriteLine($"Тип: {order.Type}");

            if (order.NoticeName != null)
                Console.WriteLine($"Title: {order.NoticeName}");

            if (order.NoticeExtension != null)
                Console.WriteLine($"Extension: {order.NoticeExtension}");

            if (order.NoticeCompany != null)
                Console.WriteLine($"NoticeCompany: {order.NoticeCompany}");

            if (order.NoticeCreator != null)
                Console.WriteLine($"Authors: {order.NoticeCreator}");

            if (order.NoticeLastAuthor != null)
                Console.WriteLine($"LastAuthor: {order.NoticeLastAuthor}");

            if (order.NoticeDateCreated != null)
                Console.WriteLine($"DateCreated: {order.NoticeDateCreated}");

            if (order.NoticeDateSaved != null)
                Console.WriteLine($"DateSaved: {order.NoticeDateSaved}");

            if (order.NoticeDatePrinted != null)
                Console.WriteLine($"DatePrinted: {order.NoticeDatePrinted}");

            if (order.NoticeTemplate != null)
                Console.WriteLine($"Template: {order.NoticeTemplate}");

            if (order.NoticeSubject != null)
                Console.WriteLine($"Subject: {order.NoticeSubject}");

            if (order.NoticeTitle != null)
                Console.WriteLine($"Title: {order.NoticeTitle}");

            Console.WriteLine();
        }

        private ExcelPackage CreateExcelFile(string resultFilePath)
        {
            FileInfo resultFileInfo = new FileInfo(resultFilePath);
            ExcelPackage excelPackage = new ExcelPackage(resultFileInfo);

            ExcelWorksheet worksheet = _excelPackage.Workbook.Worksheets.Add("Results");
            worksheet.Cells.LoadFromCollection(Enumerable.Empty<Order>(), true);

            return excelPackage;
        }

        private void WriteOrderToConsole(Order order)
        {
            Console.WriteLine();
            Console.WriteLine($"Номер закупки: {order.Number}");
            Console.WriteLine($"Размещено: {order.StartedAt.ToShortDateString()}");
            Console.WriteLine($"Обновлено: {order.UpdatedAt.ToShortDateString()}");
            Console.WriteLine($"Статус: {order.State}");
            Console.WriteLine($"Объект закупки: {order.Content}");
            Console.WriteLine($"Заказчик: {order.Customer}");
            Console.WriteLine($"Начальная цена: {order.StartPrice}");

            if (order.ResultPrice != null)
                Console.WriteLine($"Цена контракта: {order.ResultPrice}");

            if (order.Supplier != null)
                Console.WriteLine($"Поставщик: {order.Supplier}");

            Console.WriteLine($"Тип: {order.Type}");

            if (order.NoticeName != null)
                Console.WriteLine($"Title: {order.NoticeName}");

            if (order.NoticeExtension != null)
                Console.WriteLine($"Extension: {order.NoticeExtension}");

            if (order.NoticeCompany != null)
                Console.WriteLine($"NoticeCompany: {order.NoticeCompany}");

            if (order.NoticeCreator != null)
                Console.WriteLine($"Authors: {order.NoticeCreator}");

            if (order.NoticeLastAuthor != null)
                Console.WriteLine($"LastAuthor: {order.NoticeLastAuthor}");

            if (order.NoticeDateCreated != null)
                Console.WriteLine($"DateCreated: {order.NoticeDateCreated}");

            if (order.NoticeDateSaved != null)
                Console.WriteLine($"DateSaved: {order.NoticeDateSaved}");

            if (order.NoticeDatePrinted != null)
                Console.WriteLine($"DatePrinted: {order.NoticeDatePrinted}");

            if (order.NoticeTemplate != null)
                Console.WriteLine($"Template: {order.NoticeTemplate}");

            if (order.NoticeSubject != null)
                Console.WriteLine($"Subject: {order.NoticeSubject}");

            if (order.NoticeTitle != null)
                Console.WriteLine($"Title: {order.NoticeTitle}");

            Console.WriteLine();
        }

        #region IDisposable Support
        private bool isDisposed = false;
        public void Dispose()
        {
            if (!isDisposed)
            {
                _excelPackage?.Dispose();
                isDisposed = true;
            }
        }
        #endregion
    }
}
