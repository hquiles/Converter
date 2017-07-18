using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml;
using static Converter.Library;

namespace Converter
{
    public class ExcelToXml
    {
        private static List<Sheet> ListOfSheets { get; set; }
        private static List<Header> ListOfHeaders { get; set; }

        private static List<string> WorksheetNames { get; set; }

        private static string Path { get; set; }

        //Constructor
        public ExcelToXml(string path)
        {
            Path = path;
        }
        
        public static void ReadExcel()
        {

            //Iterate every worksheet
            for (int i = 0; i < GetWorksheetNames(Path).Count(); i++)
            {
                //Set Properties
                WorksheetNames = GetWorksheetNames(Path);

                //TODO: https://msdn.microsoft.com/library/bb332058.aspx
            }


        }
    }
}
