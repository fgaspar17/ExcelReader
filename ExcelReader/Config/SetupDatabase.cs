using Microsoft.Data.Sqlite;

namespace ExcelReader;

public class SetupDatabase
{
    public static void ResetDatabase(string file)
    {
        FileInfo fi = new(file);
        try
        {
            if (fi.Exists)
            {
                SqliteConnection connection = new(GlobalConfig.ConnectionString);
                connection.Close();
                GC.Collect();
                GC.WaitForPendingFinalizers();
                fi.Delete();
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine(ex.Message);
            fi.Delete();
        }
    }

    public static void CreateTable(List<string> columns)
    {
        string columnsFormatted = string.Join("", columns.Select(c => $"\"{c}\" TEXT NULL, \n").ToList());

        try
        {
            using SqliteConnection connection = new(GlobalConfig.ConnectionString);

            connection.Open();

            SqliteCommand cmd = connection.CreateCommand();
            cmd.CommandText = @$"CREATE TABLE IF NOT EXISTS ""Excel"" (
	                        ""IdDb""	INTEGER NOT NULL,
	                        {columnsFormatted}
	                        PRIMARY KEY(""IdDb"" AUTOINCREMENT)
                            );";
            cmd.ExecuteNonQuery();

            connection.Close();

        }
        catch (Exception ex)
        {

            Console.WriteLine($"An error ocurred: {ex.Message}");
        }
    }
}