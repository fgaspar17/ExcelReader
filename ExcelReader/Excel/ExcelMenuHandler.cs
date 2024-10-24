using Spectre.Console;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelReader;

public class ExcelMenuHandler
{
    private string _currentWorkSheet; 

    public ExcelMenuHandler(string currentWorkSheet)
    {
        _currentWorkSheet = currentWorkSheet;
    }
    public void Display()
    {
        MenuPresentation.MenuDisplayer<ExcelMenuOptions>(() => $"[green]Excel Menu - {_currentWorkSheet}[/]", HandleMenuOptions);
    }

    private bool HandleMenuOptions(ExcelMenuOptions option)
    {
        switch (option)
        {
            case ExcelMenuOptions.Exit:
                return false;
            case ExcelMenuOptions.SelectWorksheet:
                SelectWorkSheet();
                break;
            case ExcelMenuOptions.ShowExcel:
                ShowExcel();
                break;
            case ExcelMenuOptions.CreateRowExcel:
                CreateRowExcel();
                break;
            case ExcelMenuOptions.SaveExcel:
                SaveExcel();
                break;
            default:
                AnsiConsole.WriteLine($"Unknow option: {option}");
                break;
        }

        return true;
    }

    private void SaveExcel()
    {
        var workSheets = new List<WorkSheet>();
        ExcelReader excelReader = new ExcelReader();
        FileInfo fileInfo = new FileInfo(GlobalConfig.FilePath!);
        int workSheetCount = excelReader.GetWorkSheetsCount(fileInfo);

        for (int i = 0; i < workSheetCount; i++)
        {
            string workSheetName = excelReader.GetWorkSheetName(fileInfo, i);
            var columns = excelReader.GetColumns(fileInfo, i);

            var data = ExcelController.GetDataFromTable(workSheetName, GlobalConfig.ConnectionString!);
            workSheets.Add(new WorkSheet
            {
                Index = i,
                Name = workSheetName,
                Data = data,
            });
        }

        Excel excel = new () { WorkSheets = workSheets };

        ExcelWriter excelWriter = new(excel);
        if(excelWriter.Write())
            AnsiConsole.MarkupLine($"[green]Excel created![/]");
        else
            AnsiConsole.MarkupLine($"[red]Error creating the Excel file[/]");
        Prompter.PressKeyToContinuePrompt();
    }

    private void CreateRowExcel()
    {
        var columns = ExcelController.GetColumnsFromTable(_currentWorkSheet, GlobalConfig.ConnectionString!);
        var row = new Dictionary<string, string>();

        foreach ( var column in columns )
        {
            var prompt = new TextPrompt<string>($"Enter {column}:");
            row.Add(column, AnsiConsole.Prompt(prompt));
        }

        if(ExcelController.InsertData(_currentWorkSheet, row, GlobalConfig.ConnectionString!))
            AnsiConsole.MarkupLine($"[green]Row inserted![/]");
        else
            AnsiConsole.MarkupLine($"[red]Error inserting the row[/]");

        Prompter.PressKeyToContinuePrompt();
    }

    private void ShowExcel()
    {
        var data = ExcelController.GetDataFromTable(_currentWorkSheet, GlobalConfig.ConnectionString!);
        OutputRenderer.ShowTable(data, _currentWorkSheet);
        Prompter.PressKeyToContinuePrompt();
    }

    private void SelectWorkSheet()
    {
        MenuPresentation.PresentMenu("[green]Select WorkSheet[/]");
        var worksheets = ExcelController.GetTables(GlobalConfig.ConnectionString!);
        var worksheetsDictionary = new Dictionary<string, List<string>>()
        {
            { "WorkSheets", worksheets }
        };
        OutputRenderer.ShowTable(worksheetsDictionary, "WorkSheets");

        bool continueAsking = true;
        string userInput = string.Empty;
        do
        {
            var prompt = new TextPrompt<string>("Enter a WorkSheet");
            userInput = AnsiConsole.Prompt(prompt);

            continueAsking = !worksheets.Contains(userInput, StringComparer.InvariantCultureIgnoreCase);
            if (!continueAsking) AnsiConsole.MarkupLine("[red]Wrong input.[/]");
        } while (continueAsking);

        _currentWorkSheet = userInput;
    }
}