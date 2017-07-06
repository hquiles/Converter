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
using static FileProcessor.Library;
using System.Diagnostics;

namespace FileProcessor
{
    public partial class FileProcessor
    {
        public static void Main(string[] args)
        {
            string strOriginalPath = "";
            string strDestinationPath = "";
            string strFileName = "";
            do
            {
                try
                {
                    if (Directory.Exists(GetAppSetting("ScanningDirectory")))
                    {
                        if (Directory.GetFiles(GetAppSetting("ScanningDirectory"), "*.xlsx").Length != 0)
                        {
                            for (int intFileIndex = 0; intFileIndex < Directory.GetFiles(GetAppSetting("ScanningDirectory"), "*.xlsx").Length; intFileIndex++)
                            {
                                //Process Excel File



                                //Backup Files
                                strFileName = Path.GetFileName(Directory.GetFiles(GetAppSetting("ScanningDirectory"), "*.xlsx")[intFileIndex]);
                                strOriginalPath = Directory.GetFiles(GetAppSetting("ScanningDirectory"), "*.xlsx")[intFileIndex];
                                strDestinationPath = Path.Combine(GetAppSetting("BackupDirectory"), strFileName);
                                BackupProcessedFile(strOriginalPath, strDestinationPath);

                            }
                        }
                    }
                    else
                    {
                        List<string> validOptions = new List<string>();
                        validOptions.Add("yes");
                        validOptions.Add("no");
                        validOptions.Add("y");
                        validOptions.Add("n");

                        Console.WriteLine($"Directory not found at {GetAppSetting("ScanningDirectory")}");
                        Console.WriteLine("Would you like to create it? enter y,yes,n,no");

                        var response = Console.ReadLine().ToLower();
                        while (!validOptions.Contains(response))
                        {
                            Console.WriteLine("Valid options are y,yes,n,no ... Try again...");
                            response = Console.ReadLine().ToLower();


                        }
                        //Exit Program
                        if (response.Equals("no") || response.Equals("n"))
                        {
                            Console.WriteLine("Directory was not created, press any key to exit");
                            Console.ReadKey();
                            Environment.Exit(0);
                        }

                        else
                        {
                            Console.WriteLine($"Are you sure you want to create directory at {GetAppSetting("ScanningDirectory")}?");
                            response = Console.ReadLine().ToLower();

                            while (!validOptions.Contains(response))
                            {
                                Console.WriteLine("Valid options are y,yes,n,no ... Try again...");
                                response = Console.ReadLine().ToLower();
                            }
                            //Exit Program
                            if (response.Equals("no") || response.Equals("n"))
                            {
                                Console.WriteLine("Directory was not created, press any key to exit");
                                Console.ReadKey();
                                Environment.Exit(0);
                            }
                            else
                            {
                                //Create Directory
                                Directory.CreateDirectory($"{GetAppSetting("ScanningDirectory")}");
                                if (Directory.Exists($"{GetAppSetting("ScanningDirectory")}"))
                                {
                                    Console.WriteLine($"Directory was created at {GetAppSetting("ScanningDirectory")}");
                                }
                            }
                        }


                    }
                }
                catch (Exception ex)
                {
                    Debug.WriteLine(ex.StackTrace);
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

