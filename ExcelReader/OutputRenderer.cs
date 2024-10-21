using Spectre.Console;
using System.Reflection;
using System.Text;
using static OfficeOpenXml.ExcelErrorValue;

namespace ExcelReader;

public static class OutputRenderer
{
    private static readonly Dictionary<Type, PropertyInfo[]> PropertiesCache = new Dictionary<Type, PropertyInfo[]>();


    /// <summary>
    /// Displays a list of objects in a table format.
    /// </summary>
    public static void ShowTable(Dictionary<string,List<string>> columnValues, string title)
    {
        // Create a table
        var table = new Table();

        // Set the table border and style options
        table.Border(TableBorder.Rounded);
        table.BorderColor(Color.LightCoral);
        table.Title($"[bold yellow]{title}[/]");
        table.Caption("Generated on " + DateTime.Now.ToString("g"));


        // Add columns
        foreach (string column in columnValues.Keys)
        {
            table.AddColumn(new TableColumn(new Markup("[bold yellow]" + column + "[/]")).Centered().PadRight(2));
        }

        var firstRow = columnValues.Values.FirstOrDefault();
        if (firstRow == null) return;
        int rowCount = firstRow.Count;
        // Add rows
        for (int i = 0; i < rowCount; i++)
        {
            List<string> row = new();

            foreach (string column in columnValues.Keys)
            {
                row.Add(columnValues[column][i] ?? "N/A");
            }
            table.AddRow(row.ToArray());
        }

        AnsiConsole.Write(table);
    }
}