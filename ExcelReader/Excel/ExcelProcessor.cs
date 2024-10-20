
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
        var columns = _excelReader.GetColumns(_fileInfo);
        SetupDatabase.CreateTable(columns);

        // TODO: Probar Excel con valores nulos, filas y campos
        var columnValues = _excelReader.GetData(_fileInfo);

        var rows = new List<Dictionary<string, string>>();

        var firstRow = columnValues.Values.FirstOrDefault();
        if (firstRow == null) return;
        int rowCount = firstRow.Count;

        for (int i = 0; i < rowCount; i++)
        {
            var row = new Dictionary<string, string>();
            foreach (var column in columns)
            { 
                row.Add(column, columnValues[column][i]);
            }
            rows.Add(row);
        }

        // TODO: Check empty rows
        foreach (var row in rows)
        {
            ExcelController.InsertData(row, GlobalConfig.ConnectionString);
        }
        
        // TODO: App
        //InsertData();
        //ReadData();
    }
}