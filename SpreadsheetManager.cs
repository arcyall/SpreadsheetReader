using Microsoft.Data.Sqlite;
using NPOI.OOXML;
using NPOI.OpenXmlFormats;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SpreadsheetReader;

public sealed class SpreadsheetManager
{
    private readonly string _path;
    private readonly string _name;
    private readonly string _filetype;
    
    public SpreadsheetManager(string path)
    {
        _path = path;
        _name = Path.GetFileName(path);
        _filetype = Path.GetExtension(path);
    }
    public void ParseFile()
    {
        Console.WriteLine($"Provided filename: {_name}");
        Console.WriteLine($"Provided filetype: {_filetype}");

        switch (_filetype)
        {
            //TODO: support more file types such as csv, ods, xls
            case ".xlsx":
                Console.WriteLine("Processing excel file");
                ProcessSheet();
                Console.WriteLine("File fully loaded into db");
                break;

            default:
                Console.WriteLine("Unsupported filetype");
                break;
        }
    }
    
    private void ProcessSheet()
    {
        // TODO: support multiple sheets per workbook
        using var stream = new FileStream(_path, FileMode.Open);
        var workbook = new XSSFWorkbook(stream);
        var sheet = workbook.GetSheetAt(0);
        var headerRow = sheet.GetRow(0);
        
        for (int i = 0; i < headerRow.LastCellNum; ++i)
        {
            var cell = headerRow.GetCell(i);
            Console.WriteLine(cell.ToString());
        }

        //using var connection = new SqliteConnection($"Data Source={_name}.db");
        //connection.Open();

        throw new NotImplementedException();
    }
}
