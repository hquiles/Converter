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
    public partial class Converter
    {

        public static string strOriginalPath { get; set; }
        public static string strDestinationPath { get; set; }
        public static string strFileName { get; set; }

        public static void Main(string[] args)
        {
            //Instantiate ExcelToXmlClass
           

            do
            {
                //Create Directories if they dont exist.
                CreateDirectories(GetAppSetting("ScanningDirectory"), GetAppSetting("BackupDirectory"));
                try
                {
                    if (Directory.GetFiles(GetAppSetting("ScanningDirectory"), "*.xlsx").Length != 0)
                    {
                        for (int intFileIndex = 0; intFileIndex < Directory.GetFiles(GetAppSetting("ScanningDirectory"), "*.xlsx").Length; intFileIndex++)
                        {
                            //Set Properties
                            strFileName = Path.GetFileName(Directory.GetFiles(GetAppSetting("ScanningDirectory"), "*.xlsx")[intFileIndex]);
                            strOriginalPath = Directory.GetFiles(GetAppSetting("ScanningDirectory"), "*.xlsx")[intFileIndex];
                            strDestinationPath = Path.Combine(GetAppSetting("BackupDirectory"), strFileName);

                            //Process Excel Files
                            new ExcelToXml(strOriginalPath);
                            ExcelToXml.ReadExcel();



                            //Backup Files                           
                            BackupProcessedFile(strOriginalPath, strDestinationPath);

                        }
                    }
                }

                catch (DirectoryNotFoundException ex)
                {
                    Console.WriteLine($"MESSAGE: { ex.Message}"); 
                    CreateDirectories(GetAppSetting("ScanningDirectory"), GetAppSetting("BackupDirectory"));
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"MESSAGE: { ex.Message}");
                }

            } while (true);
        }
    }
}

