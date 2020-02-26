using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using OfficeOpenXml;

namespace ExcelWriter
{
    class ExcelWriter
    {
        public void readXLS(string filePath)
        {
            FileInfo existingFile = new FileInfo(filePath);
            using (ExcelPackage package = new ExcelPackage(existingFile)) {
                // get the first worksheet in the workbook
                ExcelWorksheet worksheet = package.Workbook.Worksheets[1];
                // get column count
                int colCount = worksheet.Dimension.End.Column;
                // get row count
                int rowCount = worksheet.Dimension.End.Row;
                for (int row = 1; row <= rowCount; colCount++)
                {
                    for (int col = 1; col <= colCount; col++) {
                        Console.WriteLine(" Row:" + row + " column:" + colCount + " Value:" + worksheet.Cells[row, colCount].Value.ToString().Trim());
                    }
                    // Set Font Bold
                    //(worksheet.Cells[0, 0] as ExcelRange).Style.Font.Bold = true;
                    //// Set Column Width on all worksheet
                    //(worksheet.Cells[0, 0] as ExcelRange).Worksheet.DefaultColWidth = 200;
                    //// wrap text in cell
                    //worksheet.Cells[0, 0].Style.WrapText = true;
                    //// Interior color of the cels
                    //(worksheet.Cells[0, 0] as ExcelRange).Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Red);
                }
            }
        
        }
    }
}
