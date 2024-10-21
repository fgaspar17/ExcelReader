
using OfficeOpenXml;
using System.IO;

namespace ExcelReader;

public class ExcelProcessor
{
    private FileInfo _fileInfo;
    private ExcelReader _excelReader;

    public ExcelProcessor(string excelPath)
    {
        _fileInfo = new FileInfo(excelPath);
        _excelReader = new ExcelReader();
    }

    public void Run()
    {
        int worksheetsCount = _excelReader.GetWorkSheetsCount(_fileInfo);
        for (int i = 0; i < worksheetsCount; i++)
        {
            var columns = _excelReader.GetColumns(_fileInfo, i);
            string worksheetName = _excelReader.GetWorkSheetName(_fileInfo, i);
            SetupDatabase.CreateTable(worksheetName, columns);

            // TODO: Probar Excel con valores nulos, filas y campos
            var columnValues = _excelReader.GetData(_fileInfo, i);

            var rows = new List<Dictionary<string, string>>();

            var firstRow = columnValues.Values.FirstOrDefault();
            if (firstRow == null) return;
            int rowCount = firstRow.Count;

            for (int r = 0; r < rowCount; r++)
            {
                var row = new Dictionary<string, string>();
                foreach (var column in columns)
                { 
                    row.Add(column, columnValues[column][r]);
                }
                rows.Add(row);
            }

            // TODO: Check empty rows
            foreach (var row in rows)
            {
                ExcelController.InsertData(worksheetName, row, GlobalConfig.ConnectionString);
            }
        }
        
        // TODO: App
        //ReadData();
    }
}