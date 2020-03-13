using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using OfficeOpenXml;

namespace ExcelWriter
{
    class ExcelWriter
    {
        ExcelWorkbook wb;

        public void readXLS(string filePath)
        {
            // get the path
            string path = Path.GetFullPath(Path.Combine(Directory.GetCurrentDirectory(), @"..\..\..\"))+ @"ExcelTemplate\Template.xlsx";

            // setting the path
            FileInfo existingFile = new FileInfo(path);

            // setting the package manager for excel
            using (ExcelPackage package = new ExcelPackage(existingFile)) {
                // verify if the path exist
                if (File.Exists(path))
                {
                    Console.WriteLine("exist");
                }
                // get the first worksheet in the workbook
                ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
                // get column count
                //int colCount = worksheet.Dimension.End.Column;
                // get row count
                //int rowCount = worksheet.Dimension.End.Row;
                //for (int row = 1; row <= rowCount; colCount++)
                //{
                //    for (int col = 1; col <= colCount; col++) {
                //        Console.WriteLine(" Row:" + row + " column:" + colCount + " Value:" + worksheet.Cells[row, colCount].Value.ToString().Trim());
                //    }
                //    // Set Font Bold
                //    //(worksheet.Cells[0, 0] as ExcelRange).Style.Font.Bold = true;
                //    //// Set Column Width on all worksheet
                //    //(worksheet.Cells[0, 0] as ExcelRange).Worksheet.DefaultColWidth = 200;
                //    //// wrap text in cell
                //    //worksheet.Cells[0, 0].Style.WrapText = true;
                //    //// Interior color of the cels
                //    //(worksheet.Cells[0, 0] as ExcelRange).Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Red);
                //}

                //Console.WriteLine(worksheet.Cells[1, 1].Value.ToString().Trim());

                // change the first column
                (worksheet.Cells["B1"] as ExcelRange).Value = "YESSS!!";
                string whereToSave = @"C:\Temp\test2.xls";
                
                FileInfo fi = new FileInfo(whereToSave);
                package.SaveAs(fi);

                // for seving from using directive
                //System.IO.File.WriteAllBytes(@"c:\Temp\myworkbook.xlsx", package);



                // Other save method
                //excel.Save(@"C:\Temp\test1.xls");
                //FileInfo fi = new FileInfo(@"C:\Temp\" + fileName);
                //package.SaveAs(fi);

                //var filePath = @"C:\Temp\"+fileName;
                //var mimeType = "application/vnd.ms-excel";
                //File(new FileStream(filePath, FileMode.Open), mimeType, fileName);

                //// Convert the excel package to a byte array
                //byte[] bin = package.GetAsByteArray();

                //// create a SeveFileDialog with some properties
                //SaveFileDialog                
                //package.Save();
                //System.IO.File.WriteAllBytes(@"c:\Temp\myworkbook.xlsx", package);

                //package.SaveAs(new FileInfo(@"c:\workbooks\myworkbook.xlsx"));

                // send email
                //new SMTPEmailManager().SendEmail(new string[] { "luca.perticaroli@dev4side.com" }, new string[0], subject, body, new MemoryStream(package), fileName);

                // save the change
                package.Save();

                // TODO: Al posto del mio template di esempioinserire quello inviato da Mattia
                // selezionare il Woprksheet corretto in base al tipo di richiesta (visit request, visit request vip, ...)
                // Inserire nei campi corrispettivi i valori
            }  
        }
    }
}
