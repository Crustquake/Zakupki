﻿using System;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Text;

using CsvHelper;

namespace Zakupki
{
    public class Program
    {
        static void Main(string[] args)
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
                    Process.Start(tempFolder);

                    string logFile = Path.Combine(tempFolder, "result.csv");

                    using (StreamWriter streamWriter = new StreamWriter(logFile, false, Encoding.UTF8))
                    using (CsvWriter csvWriter = new CsvWriter(streamWriter, CultureInfo.InvariantCulture))
                    {
                        csvWriter.Configuration.Delimiter = ";";
                        csvWriter.WriteHeader(typeof(Order));
                        csvWriter.NextRecord();

                        foreach (Order order in orders)
                        {
                            if (order != null)
                            {
                                service.GetOrderResult(order);
                                var orderNoticeFile = service.DownloadOrderNotice(order, tempFolder);
                                if (orderNoticeFile != null)
                                {
                                    service.GetOrderNoticeDetails(order, orderNoticeFile);
                                }

                                WriteOrderToConsole(order);

                                csvWriter.WriteRecord(order);
                                csvWriter.NextRecord();
                            }
                        }
                    }
                }
            }
            Console.WriteLine("Приложение закончило работу, нажмите любую кнопку чтобы выйти...");
            Console.ReadKey();
        }

        private static void WriteOrderToConsole(Order order)
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
    }
}
