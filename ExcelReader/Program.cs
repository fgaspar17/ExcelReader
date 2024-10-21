// See https://aka.ms/new-console-template for more information
using ExcelReader;
using OfficeOpenXml;

ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

Startup.StartApplication();

ExcelProcessor excelProcessor = new ExcelProcessor("C:\\Users\\ErNaN\\Downloads\\SaleData(1).xlsx");
excelProcessor.Run();


//using (ExcelPackage package = new ExcelPackage(fileInfo))
//{
//    // Get the first worksheet in the workbook
//    ExcelWorksheet worksheet = package.Workbook.Worksheets[0];

//    // Get the dimensions of the worksheet
//    int rowCount = worksheet.Dimension.Rows;
//    int colCount = worksheet.Dimension.Columns;

//    // Loop through rows and columns and print data to the console
//    for (int row = 1; row <= rowCount; row++)
//    {
//        for (int col = 1; col <= colCount; col++)
//        {
//            // Print cell value to console
//            Console.Write(worksheet.Cells[row, col].Text + "\t");
//        }
//        Console.WriteLine(); // New line after each row
//    }
//}