using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.IO;
using System.Configuration;
using System.Collections.Specialized;
using static Converter.Library;
using System.Diagnostics;

namespace Converter
{
    public partial class ExcelConverter
    {
        public static void Main(string[] args)
        {
            string strDirectory = "";
            string strCurrentFilePath = "";
            string[] strArrFiles = null;
            string strCurrentFileName = "";
            string strCurrentExtension = "";


            do
            {
                try
                {
                    if (Directory.Exists(GetAppSetting("ScanningDirectory")))
                    {
                        strDirectory = GetAppSetting("ScanningDirectory");

                        if (Directory.GetFiles(strDirectory, "*.xlsx").Length != 0)
                        {
                            strArrFiles = Directory.GetFiles(strDirectory, "*.xlsx");

                            for (int file = 0; file < strArrFiles.Length; file++)
                            {
                                strCurrentFilePath = strArrFiles[file];
                                strCurrentFileName = Path.GetFileName(strCurrentFilePath);
                                strCurrentExtension = Path.GetExtension(strCurrentFilePath);

                                SetDocumentProperty(strCurrentFilePath, DocProperty.Title, strCurrentFileName);                               

                                Debug.WriteLine($"Processing {file + 1} file.");
                                Debug.WriteLine($"{Path.GetFileName(strCurrentFilePath)}");
                                Debug.WriteLine("");

                                File.Move(strArrFiles[file], $"{GetAppSetting("BackupDirectory")}{"\\"}" +
                                    $"{Guid.NewGuid().ToString()}{strCurrentExtension}");

                                System.Threading.Thread.Sleep(1000);                               

                            }
                        }
                    }

                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"{ex.InnerException}\n{ex.StackTrace.Trim()}\n{strCurrentFileName}");
                    throw;
                }
            } while (true);




            //using (SpreadsheetDocument spreadSheetDocument = SpreadsheetDocument.Open(@"D:\my desktop\temp\GoogleTests.xlsx", false))
            //{
            //    WorkbookPart wbPart = spreadSheetDocument.WorkbookPart;



            //    OpenXmlReader reader = OpenXmlReader.Create(wbPart);

            //    while (reader.Read())
            //    {
            //        if (reader.ElementType == typeof(Sheet))
            //        {
            //            Sheet sheet = (Sheet)reader.LoadCurrentElement();

            //            WorksheetPart wsPart = (WorksheetPart)(wbPart.GetPartById(sheet.Id));

            //            OpenXmlReader wsReader = OpenXmlReader.Create(wsPart);
            //            while (wsReader.Read())
            //            {
            //                if (wsReader.ElementType == typeof(Worksheet))
            //                {
            //                    Worksheet wsPartXml = (Worksheet)wsReader.LoadCurrentElement();
            //                    //Console.WriteLine(wsPartXml.OuterXml + "\n");
            //                    Console.WriteLine(wsPartXml.InnerXml);
            //                }
            //            }
            //        }
            //   }


        }
    }
}

