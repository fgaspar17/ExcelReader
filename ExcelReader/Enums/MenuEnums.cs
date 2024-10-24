using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelReader;

public enum ExcelMenuOptions
{
    [Title("Exit")]
    Exit,
    [Title("Select Worksheet")]
    SelectWorksheet,
    [Title("Show the Excel from Sqlite")]
    ShowExcel,
    [Title("Create an Excel row")]
    CreateRowExcel,
    [Title("Save the Excel")]
    SaveExcel,
}