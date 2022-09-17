﻿using Microsoft.Data.Sqlite;
using Microsoft.VisualBasic;
using NPOI.SS.UserModel;
using System.Data;
using System.Text;

namespace SpreadsheetReader;

public sealed class SpreadsheetManager
{
    private readonly string _path;
    private readonly string _name;
    private readonly string _filetype;

    public SpreadsheetManager(string path)
    {
        _path = path;
        _filetype = Path.GetExtension(path);
        _name = Path.GetFileNameWithoutExtension(path);
    }
    public void ParseFile()
    {
        Console.WriteLine($"Provided filename: {_name}");
        Console.WriteLine($"Provided filetype: {_filetype}");

        switch (_filetype)
        {
            //TODO: support more file types such as csv, ods
            //TODO: db into spreadsheet support
            case ".xls":
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
        var createTableCmd = new StringBuilder($"CREATE TABLE IF NOT EXISTS {_name} (");
        var fillTableCmd = new StringBuilder($"INSERT INTO {_name} (");
        

        using (var stream = new FileStream(_path, FileMode.Open))
        {
            var workbook = WorkbookFactory.Create(stream);
            var sheet = workbook.GetSheetAt(0);
            var headerRow = sheet.GetRow(sheet.FirstRowNum);

            for (int i = 0; i < headerRow.LastCellNum; ++i)
            {
                var currHeader = headerRow.GetCell(i).ToString().Replace(' ', '_');
                var type = ParseType(sheet.GetRow(sheet.FirstRowNum + 1).GetCell(i).ToString());
                createTableCmd.Append('[').Append(currHeader).Append("] ").Append(type);
                fillTableCmd.Append(currHeader);
                if (i < headerRow.LastCellNum - 1)
                {
                    createTableCmd.Append(',');
                    fillTableCmd.Append(',');
                }
            }

            createTableCmd.Append(')');
            fillTableCmd.Append(") VALUES (");

            ExecuteDbCmd(createTableCmd.ToString());

            var fillTableString = fillTableCmd.ToString();

            for (int i = sheet.FirstRowNum + 1; i < sheet.LastRowNum; ++i)
            {
                var fillTableCmdTmp = new StringBuilder(fillTableString);
                var currRow = sheet.GetRow(i);

                for (int j = 0; j < currRow.Cells.Count; ++j)
                {
                    fillTableCmdTmp.Append(currRow.GetCell(j).ToString().Replace('-', '.'));
                    if (j < currRow.Cells.Count - 1)
                    {
                        fillTableCmdTmp.Append(',');
                    }
                }
                fillTableCmdTmp.Append(')');
                ExecuteDbCmd(fillTableCmdTmp.ToString());
            }
        }
    }
    private string ParseType(string val)
    {
        if (String.IsNullOrEmpty(val) || String.IsNullOrWhiteSpace(val)) return null;
        if (Int64.TryParse(val, out _)) return "BIGINT";
        if (Decimal.TryParse(val, out _)) return "DECIMAL";
        if (DateTime.TryParse(val, out _)) return "DATETIME";
        return "NVARCHAR(255)";
    }
    private void ExecuteDbCmd(string cmd)
    {
        using var connection = new SqliteConnection($"Data Source={_name}.db");
        connection.Open();
        using var command = connection.CreateCommand();
        command.CommandText = cmd;
        Console.WriteLine(cmd);
        command.ExecuteNonQuery();
    }
}
