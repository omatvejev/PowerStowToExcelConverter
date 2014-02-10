using System;
using System.IO;
using System.Collections.Generic;
using System.Drawing;
using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Drawing.Vml;
using OfficeOpenXml.Style;
using PowerStowToExcelConverter.Model;

namespace PowerStowToExcelConverter.Core
{
    class Writer
    {
        private ExcelPackage ep;
        private ExcelWorksheet ws;
        private Translator translator;

        // Modify in the future to tell the user that the file is in use somewhere outside of the constructor
        public Writer(string path, bool delete)
        {
            FileInfo fileInfo = new FileInfo(path);

            // Check if the file exists
            if (delete && fileInfo.Exists)
            {
                try
                {
                    fileInfo.Delete();
                }
                catch (Exception ex)
                {
                    throw ex;
                }
            }

            try
            {
                // Create the new excel package that will create the new file
                ep = new ExcelPackage(fileInfo);
                ws = ep.Workbook.Worksheets.Add("Content");

                // Zoom documnet to 90%
                ws.View.ZoomScale = 90;

                // Do not show gridlines
                ws.View.ShowGridLines = false;
            }
            catch (Exception ex)
            {
                throw ex;
            }

            // Create a translator object. In-case if there is any problem then set the object to null
            try
            {
                translator = new Translator(@"Translation.xlsx");
            }
            catch (Exception ex)
            {
                translator = null;
                throw ex;
            }
        }

        // Checks if the file exists and does appropriate actions
        public static bool fileExists(string path)
        {
            FileInfo fileInfo = new FileInfo(path);

            if (fileInfo.Exists)
                return true;
            else
                return false;
        }

        public void writeTerminalInformation(Terminal terminal, int row, int col)
        {
            ws.Cells[row, col].Style.Font.Bold = true;
            ws.Cells[row, col].Style.Font.Size = 14;
            ws.Cells[row, col].Value = terminal.Name;

            ws.Cells[row, col + 11, row, col + 13].Merge = true;
            ws.Cells[row, col + 11].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Right;
            ws.Cells[row, col + 11].Style.Font.Bold = true;
            ws.Cells[row, col + 11].Style.Font.Size = 12;
            ws.Cells[row, col + 11].Value = "Vessel Code:";

            ws.Cells[row, col + 16, row, col + 18].Merge = true;
            ws.Cells[row, col + 16].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Right;
            ws.Cells[row, col + 16].Style.Font.Bold = true;
            ws.Cells[row, col + 16].Style.Font.Size = 12;
            ws.Cells[row, col + 16].Value = "Vessel Name:";

            ws.Cells[row, col + 19].Style.Font.Bold = true;
            ws.Cells[row, col + 19].Style.Font.Size = 12;
            ws.Cells[row, col + 19].Value = terminal.Vessel;

            ws.Cells[row, col + 25, row, col + 27].Merge = true;
            ws.Cells[row, col + 25].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Right;
            ws.Cells[row, col + 25].Style.Font.Bold = true;
            ws.Cells[row, col + 25].Style.Font.Size = 12;
            ws.Cells[row, col + 25].Value = "Voyage No:";

            ws.Cells[row, col + 28].Style.Font.Bold = true;
            ws.Cells[row, col + 28].Style.Font.Size = 12;
            ws.Cells[row, col + 28].Value = terminal.Voyage;
        }

        // Write the shipment information with specified company and location. Heading paramater controls if the table heading row should be included
        public void writeShipmentInformation(ShipmentCompany company, int row, int col, bool heading)
        {
            createShipmentTemplate(company.Shipments.Count, row, col, heading);
            populateShipmentData(company, row, col);
            createShipmentFormulas(company.Shipments.Count, row, col);
        }

        // Used to save the file with any new information written to it
        public void save()
        {
            ep.Save();
        }

        // Draws the template without any data in it based on the row and column position
        private void createShipmentTemplate(int numOfPorts, int row, int col, bool heading)
        {

            // Adjust the column size. (Note: The value that excel shows in the width
            // column is not the same as the value set below. There is usually around .71 difference
            // So, suppose excel shows 4.19 then the value needs to be set to 4.90)
            for (int i = col; i <= col + 19; i++)
            {
                ws.Column(i).BestFit = false;

                // First column
                if (i == col)
                    ws.Column(i).Width = 6.14;
                // OOG column
                else if (i == col + 15)
                    ws.Column(i).Width = 5.00;
                // TEUS column
                else if (i == col + 16)
                    ws.Column(i).Width = 7.14;
                // MTY WGT column
                else if (i == col + 17)
                    ws.Column(i).Width = 5.57;
                // WGT column
                else if (i == col + 18)
                    ws.Column(i).Width = 6.71;
                // The seperator Column
                else if (i == col + 19)
                    ws.Column(i).Width = 2.14;
                else
                    ws.Column(i).Width = 4.85;
            }

            if (heading)
            {
                // First Row
                ws.Cells[row, col].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Medium;
                ws.Cells[row, col].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Medium;

                ws.Cells[row, col + 1].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Medium;
                ws.Cells[row, col + 4].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Medium;
                ws.Cells[row, col + 1, row, col + 4].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Medium;
                ws.Cells[row, col + 1, row, col + 4].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Medium;
                ws.Cells[row, col + 1, row, col + 4].Merge = true;
                ws.Cells[row, col + 1].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                ws.Cells[row, col + 1].Style.Font.Italic = true;
                ws.Cells[row, col + 1].Style.Font.Bold = true;
                ws.Cells[row, col + 1].Style.Font.Size = 9;
                ws.Cells[row, col + 1].Value = "DRY";

                ws.Cells[row, col + 5].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Medium;
                ws.Cells[row, col + 7].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Medium;
                ws.Cells[row, col + 5, row, col + 7].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Medium;
                ws.Cells[row, col + 5, row, col + 7].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Medium;
                ws.Cells[row, col + 5, row, col + 7].Merge = true;
                ws.Cells[row, col + 5].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                ws.Cells[row, col + 5].Style.Font.Italic = true;
                ws.Cells[row, col + 5].Style.Font.Bold = true;
                ws.Cells[row, col + 5].Style.Font.Size = 9;
                ws.Cells[row, col + 5].Value = "REFER";

                ws.Cells[row, col + 8].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Medium;
                ws.Cells[row, col + 11].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Medium;
                ws.Cells[row, col + 8, row, col + 11].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Medium;
                ws.Cells[row, col + 8, row, col + 11].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Medium;
                ws.Cells[row, col + 8, row, col + 11].Merge = true;
                ws.Cells[row, col + 8].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                ws.Cells[row, col + 8].Style.Font.Italic = true;
                ws.Cells[row, col + 8].Style.Font.Bold = true;
                ws.Cells[row, col + 8].Style.Font.Size = 9;
                ws.Cells[row, col + 8].Value = "EMPTY";

                ws.Cells[row, col + 12].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Medium;
                ws.Cells[row, col + 14].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Medium;
                ws.Cells[row, col + 12, row, col + 14].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Medium;
                ws.Cells[row, col + 12, row, col + 14].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Medium;
                ws.Cells[row, col + 12, row, col + 14].Merge = true;
                ws.Cells[row, col + 12].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                ws.Cells[row, col + 12].Style.Font.Italic = true;
                ws.Cells[row, col + 12].Style.Font.Bold = true;
                ws.Cells[row, col + 12].Style.Font.Size = 9;
                ws.Cells[row, col + 12].Value = "NON-MT TOTALS";

                // Second Row
                ws.Cells[row + 1, col + 1].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                ws.Cells[row + 1, col + 1].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                ws.Cells[row + 1, col + 1].Style.Font.Italic = true;
                ws.Cells[row + 1, col + 1].Style.Font.Bold = true;
                ws.Cells[row + 1, col + 1].Style.Font.Size = 10;
                ws.Cells[row + 1, col + 1].Value = "20'";

                ws.Cells[row + 1, col + 2].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                ws.Cells[row + 1, col + 2].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                ws.Cells[row + 1, col + 2].Style.Font.Italic = true;
                ws.Cells[row + 1, col + 2].Style.Font.Bold = true;
                ws.Cells[row + 1, col + 2].Style.Font.Size = 10;
                ws.Cells[row + 1, col + 2].Value = "40'";

                ws.Cells[row + 1, col + 3].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                ws.Cells[row + 1, col + 3].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                ws.Cells[row + 1, col + 3].Style.Font.Italic = true;
                ws.Cells[row + 1, col + 3].Style.Font.Bold = true;
                ws.Cells[row + 1, col + 3].Style.Font.Size = 10;
                ws.Cells[row + 1, col + 3].Value = "HC";

                ws.Cells[row + 1, col + 4].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Medium;
                ws.Cells[row + 1, col + 4].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                ws.Cells[row + 1, col + 4].Style.Font.Italic = true;
                ws.Cells[row + 1, col + 4].Style.Font.Bold = true;
                ws.Cells[row + 1, col + 4].Style.Font.Size = 10;
                ws.Cells[row + 1, col + 4].Value = "45'";

                ws.Cells[row + 1, col + 5].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                ws.Cells[row + 1, col + 5].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                ws.Cells[row + 1, col + 5].Style.Font.Italic = true;
                ws.Cells[row + 1, col + 5].Style.Font.Bold = true;
                ws.Cells[row + 1, col + 5].Style.Font.Size = 10;
                ws.Cells[row + 1, col + 5].Value = "20'";

                ws.Cells[row + 1, col + 6].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                ws.Cells[row + 1, col + 6].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                ws.Cells[row + 1, col + 6].Style.Font.Italic = true;
                ws.Cells[row + 1, col + 6].Style.Font.Bold = true;
                ws.Cells[row + 1, col + 6].Style.Font.Size = 10;
                ws.Cells[row + 1, col + 6].Value = "40'";

                ws.Cells[row + 1, col + 7].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Medium;
                ws.Cells[row + 1, col + 7].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                ws.Cells[row + 1, col + 7].Style.Font.Italic = true;
                ws.Cells[row + 1, col + 7].Style.Font.Bold = true;
                ws.Cells[row + 1, col + 7].Style.Font.Size = 10;
                ws.Cells[row + 1, col + 7].Value = "HC";

                ws.Cells[row + 1, col + 8].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                ws.Cells[row + 1, col + 8].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                ws.Cells[row + 1, col + 8].Style.Font.Italic = true;
                ws.Cells[row + 1, col + 8].Style.Font.Bold = true;
                ws.Cells[row + 1, col + 8].Style.Font.Size = 10;
                ws.Cells[row + 1, col + 8].Value = "20'";

                ws.Cells[row + 1, col + 9].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                ws.Cells[row + 1, col + 9].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                ws.Cells[row + 1, col + 9].Style.Font.Italic = true;
                ws.Cells[row + 1, col + 9].Style.Font.Bold = true;
                ws.Cells[row + 1, col + 9].Style.Font.Size = 10;
                ws.Cells[row + 1, col + 9].Value = "40'";

                ws.Cells[row + 1, col + 10].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                ws.Cells[row + 1, col + 10].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                ws.Cells[row + 1, col + 10].Style.Font.Italic = true;
                ws.Cells[row + 1, col + 10].Style.Font.Bold = true;
                ws.Cells[row + 1, col + 10].Style.Font.Size = 10;
                ws.Cells[row + 1, col + 10].Value = "HC";

                ws.Cells[row + 1, col + 11].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Medium;
                ws.Cells[row + 1, col + 11].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                ws.Cells[row + 1, col + 11].Style.Font.Italic = true;
                ws.Cells[row + 1, col + 11].Style.Font.Bold = true;
                ws.Cells[row + 1, col + 11].Style.Font.Size = 10;
                ws.Cells[row + 1, col + 11].Value = "45'";

                ws.Cells[row + 1, col + 12].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                ws.Cells[row + 1, col + 12].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                ws.Cells[row + 1, col + 12].Style.Font.Italic = true;
                ws.Cells[row + 1, col + 12].Style.Font.Bold = true;
                ws.Cells[row + 1, col + 12].Style.Font.Size = 10;
                ws.Cells[row + 1, col + 12].Value = "20'";

                ws.Cells[row + 1, col + 13].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                ws.Cells[row + 1, col + 13].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                ws.Cells[row + 1, col + 13].Style.Font.Italic = true;
                ws.Cells[row + 1, col + 13].Style.Font.Bold = true;
                ws.Cells[row + 1, col + 13].Style.Font.Size = 10;
                ws.Cells[row + 1, col + 13].Value = "40'";

                ws.Cells[row + 1, col + 14].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Medium;
                ws.Cells[row + 1, col + 14].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                ws.Cells[row + 1, col + 14].Style.Font.Italic = true;
                ws.Cells[row + 1, col + 14].Style.Font.Bold = true;
                ws.Cells[row + 1, col + 14].Style.Font.Size = 10;
                ws.Cells[row + 1, col + 14].Value = "45'";

                ws.Cells[row + 1, col + 15].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Medium;
                ws.Cells[row + 1, col + 15].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Medium;
                ws.Cells[row + 1, col + 15].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                ws.Cells[row + 1, col + 15].Style.Font.Italic = true;
                ws.Cells[row + 1, col + 15].Style.Font.Bold = true;
                ws.Cells[row + 1, col + 15].Style.Font.Size = 10;
                ws.Cells[row + 1, col + 15].Value = "OOG";

                ws.Cells[row + 1, col + 16].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Medium;
                ws.Cells[row + 1, col + 16].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Medium;
                ws.Cells[row + 1, col + 16].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                ws.Cells[row + 1, col + 16].Style.Font.Italic = true;
                ws.Cells[row + 1, col + 16].Style.Font.Bold = true;
                ws.Cells[row + 1, col + 16].Style.Font.Size = 10;
                ws.Cells[row + 1, col + 16].Value = "TEUS";

                ws.Cells[row + 1, col + 17].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Medium;
                ws.Cells[row + 1, col + 17].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Medium;
                ws.Cells[row + 1, col + 17].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                ws.Cells[row + 1, col + 17].Style.WrapText = true;
                ws.Cells[row + 1, col + 17].Style.Font.Italic = true;
                ws.Cells[row + 1, col + 17].Style.Font.Bold = true;
                ws.Cells[row + 1, col + 17].Style.Font.Size = 10;
                ws.Cells[row + 1, col + 17].Value = "MTY WGT";

                ws.Cells[row + 1, col + 18].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Medium;
                ws.Cells[row + 1, col + 18].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Medium;
                ws.Cells[row + 1, col + 18].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                ws.Cells[row + 1, col + 18].Style.Font.Italic = true;
                ws.Cells[row + 1, col + 18].Style.Font.Bold = true;
                ws.Cells[row + 1, col + 18].Style.Font.Size = 10;
                ws.Cells[row + 1, col + 18].Value = "WGT";
            }

            // Company name is always included even if headings are disabled
            ws.Row(row + 1).Height = 26.25; // We need to set the hight manually in the next shipment rows
            ws.Cells[row + 1, col].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Medium;
            ws.Cells[row + 1, col].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Medium;
            ws.Cells[row + 1, col].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Medium;
            ws.Cells[row + 1, col].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Medium;
            ws.Cells[row + 1, col].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
            ws.Cells[row + 1, col].Style.Font.Bold = true;
            ws.Cells[row + 1, col].Style.Font.Size = 12;

            // Create colors
            Color yellow = ColorTranslator.FromHtml("#ffff99");
            Color turquoise = ColorTranslator.FromHtml("#69ffff");
            Color lightGreen = ColorTranslator.FromHtml("#ccffcc");
            Color lightGrey = ColorTranslator.FromHtml("#e3e3e3");
            Color lightBlue = ColorTranslator.FromHtml("#a0e0e0");
            Color grey = ColorTranslator.FromHtml("#c0c0c0");

            // Populate port rows
            for (int i = 0; i < numOfPorts; i++)
            {
                // Set up the text for the first column that has the port name in it
                ws.Cells[row + 2 + i, col].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Right;
                ws.Cells[row + 2 + i, col].Style.Font.Italic = true;

                // Set up the text for all the future container numbers
                ws.Cells[row + 2 + i, col, row + 2 + i, col + 18].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                ws.Cells[row + 2 + i, col, row + 2 + i, col + 18].Style.Font.Bold = true;
                ws.Cells[row + 2 + i, col, row + 2 + i, col + 18].Style.Font.Size = 10;

                // Third row has a medium border on the top
                if (i == 0)
                {
                    ws.Cells[row + 2, col, row + 2, col + 18].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Medium;
                }
                else
                {
                    ws.Cells[row + 2 + i, col, row + 2 + i, col + 18].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                }

                // First column has a medium left and right border
                ws.Cells[row + 2 + i, col].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Medium;
                ws.Cells[row + 2 + i, col].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Medium;

                // Dry columns
                ws.Cells[row + 2 + i, col + 1, row + 2 + i, col + 3].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                ws.Cells[row + 2 + i, col + 1, row + 2 + i, col + 4].Style.Fill.PatternType = ExcelFillStyle.Solid;
                ws.Cells[row + 2 + i, col + 1, row + 2 + i, col + 4].Style.Fill.BackgroundColor.SetColor(yellow); 
                ws.Cells[row + 2 + i, col + 4].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Medium;

                // Refer columns
                ws.Cells[row + 2 + i, col + 5, row + 2 + i, col + 6].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                ws.Cells[row + 2 + i, col + 5, row + 2 + i, col + 7].Style.Fill.PatternType = ExcelFillStyle.Solid;
                ws.Cells[row + 2 + i, col + 5, row + 2 + i, col + 7].Style.Fill.BackgroundColor.SetColor(turquoise); 
                ws.Cells[row + 2 + i, col + 7].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Medium;

                // Empty columns
                ws.Cells[row + 2 + i, col + 8, row + 2 + i, col + 10].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                ws.Cells[row + 2 + i, col + 8, row + 2 + i, col + 11].Style.Fill.PatternType = ExcelFillStyle.Solid;
                ws.Cells[row + 2 + i, col + 8, row + 2 + i, col + 11].Style.Fill.BackgroundColor.SetColor(lightGreen); 
                ws.Cells[row + 2 + i, col + 11].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Medium;

                // Non-MT columns
                ws.Cells[row + 2 + i, col + 12, row + 2 + i, col + 13].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                ws.Cells[row + 2 + i, col + 12, row + 2 + i, col + 14].Style.Fill.PatternType = ExcelFillStyle.Solid;
                ws.Cells[row + 2 + i, col + 12, row + 2 + i, col + 14].Style.Fill.BackgroundColor.SetColor(lightGrey); 
                ws.Cells[row + 2 + i, col + 14].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Medium;

                // OOG - WGT columns border
                ws.Cells[row + 2 + i, col + 15, row + 2 + i, col + 18].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Medium;

                // OOG
                ws.Cells[row + 2 + i, col + 15].Style.Fill.PatternType = ExcelFillStyle.Solid;
                ws.Cells[row + 2 + i, col + 15].Style.Fill.BackgroundColor.SetColor(lightBlue);

                // MTY WGT
                ws.Cells[row + 2 + i, col + 17].Style.Fill.PatternType = ExcelFillStyle.Solid;
                ws.Cells[row + 2 + i, col + 17].Style.Fill.BackgroundColor.SetColor(lightGreen);

                // WGT
                ws.Cells[row + 2 + i, col + 18].Style.Fill.PatternType = ExcelFillStyle.Solid;
                ws.Cells[row + 2 + i, col + 18].Style.Fill.BackgroundColor.SetColor(grey);

            }

            // Last row
            row = row + 2 + numOfPorts; // Move to the last row

            // Add the top and bottom border to the whole row
            ws.Cells[row, col, row, col + 18].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Medium;
            ws.Cells[row, col, row, col + 18].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Medium;

            // Add the text style across the row
            ws.Cells[row, col, row, col + 18].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
            ws.Cells[row, col, row, col + 18].Style.Font.Bold = true;
            ws.Cells[row, col, row, col + 18].Style.Font.Size = 10;

            // First column has a medium right and left border
            ws.Cells[row, col].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Medium;
            ws.Cells[row, col].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Medium;
            ws.Cells[row, col].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Right;
            ws.Cells[row, col].Style.Font.Italic = true;
            ws.Cells[row, col].Value = "TOT";

            // Dry columns
            ws.Cells[row, col + 1, row, col + 3].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            ws.Cells[row, col + 4].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Medium;

            // Refer columns
            ws.Cells[row, col + 5, row, col + 6].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            ws.Cells[row, col + 7].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Medium;

            // Empty columns
            ws.Cells[row, col + 8, row, col + 10].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            ws.Cells[row, col + 11].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Medium;

            // Non-MT columns
            ws.Cells[row, col + 12, row, col + 13].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            ws.Cells[row, col + 14].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Medium;

            // OG - WGT columns
            ws.Cells[row, col + 15, row, col + 18].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Medium;
        }

        // Populate the data with the provided value based on the row and column position
        private void populateShipmentData(ShipmentCompany company, int row, int col)
        {
            // Shipment Company name
            ws.Cells[row + 1, col].Value = company.Name + "\n";

            int i = 0;
            foreach (ShipmentData shipment in company.Shipments)
            { 
                // Full containers
                ws.Cells[row + 2 + i, col].Value = portInitials(shipment.Location);

                if (shipment.FullContainers.twenty != 0)
                {
                    ws.Cells[row + 2 + i, col + 1].Value = shipment.FullContainers.twenty;
                }

                if (shipment.FullContainers.forty != 0)
                {
                    ws.Cells[row + 2 + i, col + 2].Value = shipment.FullContainers.forty;
                }

                if (shipment.FullContainers.fortyHC != 0)
                {
                    ws.Cells[row + 2 + i, col + 3].Value = shipment.FullContainers.fortyHC;
                }

                if (shipment.FullContainers.fortyFive != 0)
                {
                    ws.Cells[row + 2 + i, col + 4].Value = shipment.FullContainers.fortyFive;
                }

                // Empty containers
                if (shipment.EmptyContainers.twenty != 0)
                {
                    ws.Cells[row + 2 + i, col + 8].Value = shipment.EmptyContainers.twenty;
                }

                if (shipment.EmptyContainers.forty != 0)
                {
                    ws.Cells[row + 2 + i, col + 9].Value = shipment.EmptyContainers.forty;
                }

                if (shipment.EmptyContainers.fortyHC != 0)
                {
                    ws.Cells[row + 2 + i, col + 10].Value = shipment.EmptyContainers.fortyHC;
                }

                if (shipment.EmptyContainers.fortyFive != 0)
                {
                    ws.Cells[row + 2 + i, col + 11].Value = shipment.EmptyContainers.fortyFive;
                }

                // Weight related values
                if (shipment.EmptyContainers.totalWeight != 0.0)
                {
                    ws.Cells[row + 2 + i, col + 17].Value = shipment.EmptyContainers.totalWeight;
                }

                if (shipment.FullContainers.totalWeight != 0.0)
                {
                    ws.Cells[row + 2 + i, col + 18].Value = shipment.FullContainers.totalWeight;
                }
                i++;
            }
        }

        private void createShipmentFormulas(int numOfPorts, int row, int col)
        {
            string formula = "";
            for (int i = 0; i < numOfPorts; i++)
            {
                // Non-MT totals for 20'
                formula = string.Format("IF({0} + {1} = 0, \"\", {0} + {1})", ws.Cells[row + 2 + i, col + 1].Address, ws.Cells[row + 2 + i, col + 5].Address);
                ws.Cells[row + 2 + i, col + 12].Formula = formula;

                // Non-MT totals for 40'
                formula = string.Format("IF({0} + {1} + {2} + {3} = 0, \"\", {0} + {1} + {2} + {3} )", ws.Cells[row + 2 + i, col + 2].Address, ws.Cells[row + 2 + i, col + 3].Address,
                                        ws.Cells[row + 2 + i, col + 6].Address, ws.Cells[row + 2 + i, col + 7].Address);
                ws.Cells[row + 2 + i, col + 13].Formula = formula;

                // Non-MT totals for 45'
                formula = string.Format("IF({0} + {1} = 0, \"\", {0} + {1})", ws.Cells[row + 2 + i, col + 4].Address, ws.Cells[row + 2 + i, col + 8].Address);
                ws.Cells[row + 2 + i, col + 14].Formula = formula;

                // TEUS
                formula = string.Format("IF(({0} + {1} + {2}) + (({3} + {4} + {5})*2) + (({6} + {7} + {8} + {9} + {10})*2) + {11} = 0, \"\", ({0} + {1} + {2}) + (({3} + {4} + {5})*2) + (({6} + {7} + {8} + {9} + {10})*2) + {11})", 
                          ws.Cells[row + 2 + i, col + 1].Address, ws.Cells[row + 2 + i, col + 5].Address, ws.Cells[row + 2 + i, col + 8].Address, ws.Cells[row + 2 + i, col + 2].Address, ws.Cells[row + 2 + i, col + 6].Address,
                          ws.Cells[row + 2 + i, col + 9].Address, ws.Cells[row + 2 + i, col + 3].Address, ws.Cells[row + 2 + i, col + 4].Address, ws.Cells[row + 2 + i, col + 7].Address, ws.Cells[row + 2 + i, col + 10].Address, 
                          ws.Cells[row + 2 + i, col + 11].Address, ws.Cells[row + 2 + i, col + 15].Address);
                ws.Cells[row + 2 + i, col + 16].Formula = formula;
            }

            string calcStartAddress = "";
            string calcEndAddress = "";

            // Creates the total row formulas
            for (int i = 1; i < 19; i++)
            {
                calcStartAddress = ws.Cells[row + 2, col + i].Address;
                calcEndAddress = ws.Cells[row + 2 + numOfPorts - 1, col + i].Address;
                ws.Cells[row + 2 + numOfPorts, col + i].Formula = string.Format("IF(SUM({0}:{1}) = 0, \"\" , SUM({0}:{1}))", calcStartAddress, calcEndAddress);
                ws.Cells[row + 2 + numOfPorts, col + i].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Medium;
            }
        }
        // Removes all unecessory data from the port name object and returns the output in human read form
        // in most cases only initials will be returned unless there is some initial that is not known then
        // the port name is returned instead.
        private string portInitials(string name)
        {
            string portName = "";

            // Split the string by the character ',' if it exists. In some cases the port name
            // might by itself
            string[] temp = name.Split(',');

            // Take only the port name and avoid any data about the country
            portName = temp[0];

            // Remove all the spaces at the end if they exist
            portName = portName.TrimEnd(' ');

            // Try to translate the appropriate 
            if (translator != null)
                return translator.translate(portName);
            else
                return portName;
        }
    }
}
