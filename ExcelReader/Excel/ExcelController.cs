using Dapper;

namespace ExcelReader;

public class ExcelController
{
    public static bool InsertData(string tableName, Dictionary<string, string> columnValue, string connStr)
    {
        var dictionary = new Dictionary<string, object>();
        foreach (var column in columnValue)
        {
            dictionary.Add($"@{column.Key}", column.Value);
        }

        var parameters = new DynamicParameters(dictionary);

        string sql = $@"INSERT INTO ""{tableName}"" ({string.Join(", ", columnValue.Keys)}) 
                                        VALUES ({string.Join(", ", columnValue.Keys.Select(c => $"@{c}"))});";

        return SqlExecutionService.ExecuteCommand(sql, parameters, connStr);
    }
}