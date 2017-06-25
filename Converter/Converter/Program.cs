using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;



namespace Converter
{
    class Program
    {
        static void Main(string[] args)
        {
            //https://code.msdn.microsoft.com/office/How-to-convert-excel-file-7a9bb404 

            using (SpreadsheetDocument spreadSheetDocument = SpreadsheetDocument.Open(@"D:\my desktop\temp\GoogleTests.xlsx", false))
            {
                WorkbookPart wbPart = spreadSheetDocument.WorkbookPart;

                OpenXmlReader reader = OpenXmlReader.Create(wbPart);

                while (reader.Read())
                {
                    if (reader.ElementType == typeof(Sheet))
                    {
                        Sheet sheet = (Sheet)reader.LoadCurrentElement();

                        WorksheetPart wsPart = (WorksheetPart)(wbPart.GetPartById(sheet.Id));

                        OpenXmlReader wsReader = OpenXmlReader.Create(wsPart);
                        while (wsReader.Read())
                        {
                            if (wsReader.ElementType == typeof(Worksheet))
                            {
                                Worksheet wsPartXml = (Worksheet)wsReader.LoadCurrentElement();
                                //Console.WriteLine(wsPartXml.OuterXml + "\n");
                                Console.WriteLine(wsPartXml.InnerXml);
                            }
                        }
                    }
                }

                Console.ReadKey();
            }
        }
    }
}
