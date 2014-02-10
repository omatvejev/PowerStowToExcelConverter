using System;
using System.IO;
using System.Collections.Generic;
using System.Drawing;
using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Drawing.Vml;
using OfficeOpenXml.Style;

namespace PowerStowToExcelConverter.Core
{
    class Translator
    {
        private ExcelPackage ep;
        private ExcelWorksheet ws;
        private Dictionary<string, string> dictionary;

        public Translator(string path)
        {
            FileInfo fileInfo = new FileInfo(path);

            // Check if the file exists
            if (fileInfo.Exists)
            {
                try
                {
                    // Load the file
                    ep = new ExcelPackage(fileInfo);
                    ws = ep.Workbook.Worksheets["Content"];
                    dictionary = new Dictionary<string, string>();
                    generateDictionary();
                }
                catch (Exception ex)
                {
                    throw ex;
                }
            }
        }

        // Reads the excel file and generates appropriate dictionary
        private void generateDictionary()
        {
            string portName;
            string portInitials;
            for (int rowNum = 1; rowNum <= ws.Dimension.End.Row; rowNum++)
            {
                portName = ws.Cells[rowNum, 1].Text;
                portInitials = ws.Cells[rowNum, 2].Text;
                dictionary.Add(portName, portInitials);
            }
        }

        public string translate(string name)
        {
            if (dictionary.ContainsKey(name))
                return dictionary[name];
            else
                return name;
        }
    }
}
