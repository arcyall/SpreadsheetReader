using Microsoft.Data.Sqlite;
using NPOI.SS.UserModel;
using System.Text;

namespace SpreadsheetReader;

public sealed class SpreadsheetManager
{
    private readonly string _path;
    private readonly string _name;
    private readonly string _filetype;
    private readonly string _dest;

    public SpreadsheetManager(string path, string filename, string dest)
    {
        _path = path + filename;
        _dest = dest + '/';
        _filetype = Path.GetExtension(filename);
        _name = Path.GetFileNameWithoutExtension(filename).Trim().Replace(' ', '_');

        if (File.Exists($"{_dest}{_name}.db"))
        {
            File.Delete($"{_dest}{_name}.db");
            Console.WriteLine($"Removed existing file {_name}.db at location {_dest}");
        }
    }
    public void ParseFile()
    {
        Console.WriteLine($"Filename: {_name}");
        Console.WriteLine($"Filetype: {_filetype}");
        Console.WriteLine($"Destination: {_dest}");

        switch (_filetype)
        {
            case ".xls":
            case ".xlsx":
                Console.WriteLine("Processing spreadsheet file");
                ProcessSpreadsheet();
                break;

            default:
                Console.WriteLine("Unsupported filetype");
                break;
        }
    }

    private void ProcessSpreadsheet()
    {
        try
        {
            using var stream = new FileStream(_path, FileMode.Open);
            var workbook = WorkbookFactory.Create(stream);
            var sqlCmds = new StringBuilder[workbook.NumberOfSheets];

            Parallel.For(0, workbook.NumberOfSheets, x =>
            {
                var sheet = workbook.GetSheetAt(x);
                var headerRow = sheet.GetRow(sheet.FirstRowNum);
                var createTableCmd = new StringBuilder($"CREATE TABLE {sheet.SheetName} (");
                var fillTableCmd = new StringBuilder($"INSERT INTO {sheet.SheetName} (");

                for (int i = 0; i < headerRow.LastCellNum; ++i)
                {
                    var currHeader = headerRow.GetCell(i).ToString()!.Trim().Replace(' ', '_');
                    var type = ParseType(sheet.GetRow(sheet.FirstRowNum + 1).GetCell(i).ToString());

                    createTableCmd.Append('[').Append(currHeader).Append("] ").Append(type);
                    fillTableCmd.Append(currHeader);

                    if (i < headerRow.LastCellNum - 1)
                    {
                        createTableCmd.Append(',');
                        fillTableCmd.Append(',');
                    }
                }

                createTableCmd.AppendLine(");");
                fillTableCmd.Append(") VALUES (");
                var fillTableString = fillTableCmd.ToString();
                fillTableCmd.Clear();

                for (int i = sheet.FirstRowNum + 1; i < sheet.LastRowNum; ++i)
                {
                    fillTableCmd.Append(fillTableString);

                    var currRow = sheet.GetRow(i);

                    for (int j = 0; j < currRow.Cells.Count; ++j)
                    {
                        fillTableCmd.Append('"').Append(currRow.GetCell(j).ToString()).Append('"');
                        if (j < currRow.Cells.Count - 1)
                        {
                            fillTableCmd.Append(',');
                        }
                    }
                    fillTableCmd.AppendLine(");");
                }

                var sqlCmd = new StringBuilder();

                sqlCmd.Append(createTableCmd);
                sqlCmd.Append(fillTableCmd);

                sqlCmds[x] = sqlCmd;

                Console.WriteLine($"{sheet.SheetName} processed");
            });

            var finalSqlCmd = new StringBuilder().AppendLine("BEGIN TRANSACTION;");

            foreach (var cmd in sqlCmds)
            {
                finalSqlCmd.Append(cmd);
            }

            using var connection = new SqliteConnection($"Data Source={_dest}{_name}.db");
            connection.Open();
            using var command = connection.CreateCommand();
            command.CommandText = finalSqlCmd.AppendLine("COMMIT;").ToString();
            command.ExecuteNonQuery();

            Console.WriteLine("File fully loaded into db");
        }
        catch (AccessViolationException e)
        {
            Console.WriteLine("Error reading spreadsheet, it is most likely opened by another program.\n" + e);
        }
        catch (FileNotFoundException e)
        {
            Console.WriteLine("File not found.\n" + e);
        }
    }
    private static string ParseType(string? val)
    {
        if (String.IsNullOrEmpty(val) || String.IsNullOrWhiteSpace(val)) throw new ArgumentNullException(val, "Error: Tried to parse empty cell in spreadsheet");
        if (Int64.TryParse(val, out _)) return "BIGINT";
        if (Decimal.TryParse(val, out _)) return "DECIMAL";
        if (DateTime.TryParse(val, out _)) return "DATETIME";
        return "NVARCHAR(255)";
    }
}
