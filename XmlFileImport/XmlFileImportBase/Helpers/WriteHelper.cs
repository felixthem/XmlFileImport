using Microsoft.Office.Interop.Excel;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Threading.Tasks;
using XmlFileImportBase.Models;
using Excel = Microsoft.Office.Interop.Excel;


namespace XmlFileImportBase.Helpers
{
    public static class WriteHelper
    {
        public static string GetStringForWrite(List<item> items)
        {
            try
            {
                string vResult = $@"";
                int i = 0;
                foreach (item it in items)
                {
                    i++;
                    vResult += $@"Публикация {i} 
Название: {it.title}
Ссылка: {it.link}
Описание: {it.description}
Категория: {it.category}
Дата публикации: {it.pubDate}
";
                }

                return vResult;
            }
            catch (Exception e)
            {
                throw new Exception($@"Возникли трудности при формировании строки для записи. {e.Message}", e);
            }
        }

        public static async Task WriteWordAsync(List<item> items, string path)
        {
            try
            {
                if(Global.overwritingFile)
                {
                    File.WriteAllText(path, string.Empty);
                }
                using (FileStream fs = File.OpenWrite(path))
                {
                    Byte[] content = new UTF8Encoding(true).GetBytes(GetStringForWrite(items));
                    await fs.WriteAsync(content, 0, content.Length);
                }
            }
            catch (Exception e)
            {
                throw new Exception($@"Возникли трудности при записи данных в Word. {e.Message}", e);
            }
        }

        public static async Task WriteTxtAsync(List<item> items, string path)
        {
            try
            {
                using (StreamWriter writer = new StreamWriter(path, Global.overwritingFile))
                {
                    await writer.WriteLineAsync(GetStringForWrite(items));
                }
            }
            catch (Exception e)
            {
                throw new Exception($@"Возникли трудности при записи данных в тестовый файл. {e.Message}", e);
            }
        }

        public static void WriteExcel(List<item> items, string path)
        {

            Application xlApp = new Application();
            Workbook xlWb = null;
            try
            {
                int iLastRow = 0;

                if (!File.Exists(path))
                {
                    ExcelPackage.LicenseContext = LicenseContext.Commercial;
                    using (var pacakage = new ExcelPackage())
                    {
                        var ws = pacakage.Workbook.Worksheets.Add("Новости");
                        pacakage.SaveAs(new FileInfo(path));
                    }
                }

                xlWb = xlApp.Workbooks.Open(path);
                Worksheet xlSht = xlWb.Sheets[1];

                if (!Global.overwritingFile)
                {
                    xlSht.Columns.Clear();
                }
                else
                {
                    iLastRow = xlSht.Cells[xlSht.Rows.Count, "A"].End[Excel.XlDirection.xlUp].Row;
                }

                if (iLastRow == 0)
                {
                    iLastRow++;
                    xlSht.Cells[iLastRow, "A"].Value = "Название";
                    xlSht.Cells[iLastRow, "B"].Value = "Ссылка";
                    xlSht.Cells[iLastRow, "C"].Value = "Описание";
                    xlSht.Cells[iLastRow, "D"].Value = "Категория";
                    xlSht.Cells[iLastRow, "E"].Value = "Дата публикации";
                }

                iLastRow++; ;
                foreach (item it in items)
                {
                    xlSht.Cells[iLastRow, "A"].Value = it.title;
                    xlSht.Cells[iLastRow, "B"].Value = it.link;
                    xlSht.Cells[iLastRow, "C"].Value = it.description;
                    xlSht.Cells[iLastRow, "D"].Value = it.category;
                    xlSht.Cells[iLastRow, "E"].Value = it.pubDate;
                    iLastRow++;
                }

                xlApp.Visible = false;
                xlWb.Close(true);
                xlApp.Quit();
            }
            catch (Exception e)
            {
                xlWb.Close(true);
                xlApp.Quit();
                throw new Exception($@"Возникли трудности при записи данных в Excel. {e.Message}", e);
            }
        }
    }
}
