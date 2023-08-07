using System;
using System.IO;
using OfficeOpenXml;

namespace Enverus
{
    public class Conversion
    {
        public void ConvertExcelToCsv()
        {
            var excelFilePath = "Book1.xlsx";
            var csvFilePath = "data.csv";
            string endYear = "2021";
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using (var package = new ExcelPackage(new FileInfo(excelFilePath)))
            {
                var worksheet = package.Workbook.Worksheets[0];

                using var csvWriter = new StreamWriter(csvFilePath);

                for (int row = 1; row <= worksheet.Dimension.Rows; row++)
                {
                    var rowData = new List<string>();

                    for (int col = 1; col <= worksheet.Dimension.Columns; col++)
                    {

                        if (endYear != Convert.ToString(worksheet.Cells[row, col].Value))
                        {
                            rowData.Add(worksheet.Cells[row, col].Value?.ToString() ?? "");
                        }
                        else
                        {
                            col = worksheet.Dimension.Columns;
                            row = worksheet.Dimension.Rows;
                            break;
                        }
                    }

                    csvWriter.WriteLine(string.Join(",", rowData));
                }
            }
        }
    }
}
