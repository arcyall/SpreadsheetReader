using Microsoft.Data.Sqlite;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SpreadsheetReader;

public static class SpreadsheetManager
{
    public static void ParseFile(string input)
    {
        string filetype = input.Substring(input.LastIndexOf('.') + 1);
        string name = input[..(input.LastIndexOf('.') - 1)];
        Console.WriteLine($"Provided filename: {name}");
        Console.WriteLine($"Provided filetype: {filetype}");

        switch (filetype)
        {
            //TODO: support more file types such as csv, ods
            case "xlsx":
                Console.WriteLine("Processing excel sheet");
                using (var package = new ExcelPackage($"./{input}"))
                {
                    for (int i = 0; i < package.Workbook.Worksheets.Count; ++i)
                    {
                        Console.WriteLine($"Processing {package.Workbook.Worksheets[i]}");
                        ProcessSheet(name, package.Workbook.Worksheets[i]);
                        Console.WriteLine($"Processed {package.Workbook.Worksheets[i].Name}");
                    }
                }
                Console.WriteLine("File fully loaded into db");
                break;

            default:
                Console.WriteLine("Unsupported filetype");
                break;
        }
    }

    public static void ProcessSheet(string name, ExcelWorksheet input)
    {
        //using var connection = new SqliteConnection($"Data Source={name}.db");
        throw new NotImplementedException();
    }
}
