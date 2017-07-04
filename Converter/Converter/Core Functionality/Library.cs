using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.IO.Packaging;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Converter
{
    public static class Library
    {
        public static Sheets GetWorksheets(string strFileName)
        {
            Sheets theSheets = null;

            using (SpreadsheetDocument document =
                SpreadsheetDocument.Open(strFileName, false))
            {
                WorkbookPart wbPart = document.WorkbookPart;
                theSheets = wbPart.Workbook.Sheets;
            }
            return theSheets;
        }
        public static int GetWorksheetCount(string strFileName)
        {
            Sheets theSheets = null;

            using (SpreadsheetDocument document =
                SpreadsheetDocument.Open(strFileName, false))
            {
                WorkbookPart wbPart = document.WorkbookPart;
                theSheets = wbPart.Workbook.Sheets;
            }
            return ((ICollection)theSheets).Count;
            throw new NotImplementedException();
        }
        public static string GetAppSetting(string key)
        {
            try
            {
                return ConfigurationManager.AppSettings[key];
            }
            catch (Exception)
            {
                throw;
            }
        }
        public static string GetDocumentProperty(string strFilePath, DocProperty docProp)
        {
            string propertyValue = "";

            using (SpreadsheetDocument wkbk = SpreadsheetDocument.Open(strFilePath, true))
            {
                switch (docProp)
                {
                    case DocProperty.Creator:
                        propertyValue = wkbk.PackageProperties.Creator;
                        break;
                    case DocProperty.LastModifiedBy:
                        propertyValue = wkbk.PackageProperties.LastModifiedBy;
                        break;
                    case DocProperty.Category:
                        propertyValue = wkbk.PackageProperties.Category;
                        break;
                    case DocProperty.Description:
                        propertyValue = wkbk.PackageProperties.Description;
                        break;
                    case DocProperty.Subject:
                        propertyValue = wkbk.PackageProperties.Subject;
                        break;
                    case DocProperty.Title:
                        propertyValue = wkbk.PackageProperties.Title;
                        break;
                }
            }
            return propertyValue;
        }
        public static void SetDocumentProperty(string strFilePath, DocProperty docProp, string strValue)
        {
            using (SpreadsheetDocument wkbk = SpreadsheetDocument.Open(strFilePath, true))
            {
                try
                {
                    switch (docProp)
                    {
                        case DocProperty.Creator:
                            wkbk.PackageProperties.Creator = strValue;
                            break;
                        case DocProperty.LastModifiedBy:
                            wkbk.PackageProperties.LastModifiedBy = strValue;
                            break;
                        case DocProperty.Category:
                            wkbk.PackageProperties.Category = strValue;
                            break;
                        case DocProperty.Description:
                            wkbk.PackageProperties.Description = strValue;
                            break;
                        case DocProperty.Subject:
                            wkbk.PackageProperties.Subject = strValue;
                            break;
                        case DocProperty.Title:
                            wkbk.PackageProperties.Title = strValue;
                            break;
                        default:
                            break;
                    }
                }
                catch (Exception)
                {
                    throw;
                }
            }
        }
        public static void BackupProcessedFile(string strOriginalFilePath, string strTargetPath, bool overwriteExisting = true)
        {
            try
            {
                File.Copy(strOriginalFilePath, strTargetPath, overwriteExisting);               
                File.Delete(strOriginalFilePath);
            }
            catch (Exception ex)
            {   }
        }
    }
    public enum DocProperty
    {
        Creator,
        Created,
        Modified,
        LastModifiedBy,
        Category,
        Description,
        Subject,
        Title
    }
}

