using System;
using System.IO;
using System.Threading.Tasks;
using XmlFileImportBase.Helpers;
using XmlFileImportBase.Models;

namespace XmlFileImportApp
{
    internal class Program
    {
        static async Task Main(string[] args)
        {
            Console.WriteLine("Начало работы приложения.");
            channel ch = XmlHelper.ParseXML($@"data.xml");

            await Task.WhenAll(WriteHelper.WriteTxtAsync(ch.Items, "txtFile.txt"));
            await Task.WhenAll(WriteHelper.WriteWordAsync(ch.Items, "wordFile.doc"));

            string path = Path.Combine(Directory.GetCurrentDirectory(), "excelFile.xlsx");
            WriteHelper.WriteExcel(ch.Items, path);
        }
    }
}
