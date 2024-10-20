using OfficeOpenXml;
using System.IO;

namespace ExcelReader;

public class ExcelReader
{
    // TODO: Include param in appsettings.json
    private string _worksheet = "SaleData";
    public ExcelReader() 
    { 
        // TODO: parameter worksheet
    }

    public List<string> GetColumns(FileInfo fileInfo)
    {
        var columns = new List<string>();

        using (ExcelPackage package = new ExcelPackage(fileInfo))
        {
            // Get the first worksheet in the workbook
            ExcelWorksheet worksheet = package.Workbook.Worksheets[_worksheet];

            int row = 1;
            // Get the dimensions of the worksheet columns
            int colCount = worksheet.Dimension.Columns;

            // Loop through the columns of the first row
            for (int col = 1; col <= colCount; col++)
            {
                columns.Add(worksheet.Cells[row, col].Text);
            }
        }

        return columns;
    }

    public Dictionary<string, List<string>> GetData(FileInfo fileInfo)
    {
        var columnValues = new Dictionary<string, List<string>>();
        var columns = GetColumns(fileInfo);
        foreach (var column in columns)
        {
            columnValues.Add(column, new List<string>());
        }

        using (ExcelPackage package = new ExcelPackage(fileInfo))
        {
            // Get the first worksheet in the workbook
            ExcelWorksheet worksheet = package.Workbook.Worksheets[_worksheet];

            // Get the dimensions of the worksheet columns and rows
            int colCount = worksheet.Dimension.Columns;
            int rowCount = worksheet.Dimension.Rows;

            for(int row = 2; row <= rowCount; row++)
            {
                // Loop through the columns of the first row
                for (int col = 1; col <= colCount; col++)
                {
                    columnValues[columns[col - 1]].Add(worksheet.Cells[row, col].Text);
                }
            }
        }

        return columnValues;
    }
}