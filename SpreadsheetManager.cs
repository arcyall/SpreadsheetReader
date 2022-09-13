using NPOI.SS.UserModel;
using System.Data;

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
        using var stream = new FileStream(_path, FileMode.Open);

        var workbook = WorkbookFactory.Create(stream);

        var sheet = workbook.GetSheetAt(0);
        var headerRow = sheet.GetRow(0);

        var table = new DataTable();
        // populate datatable with columns from header row
        for (int i = 0; i < headerRow.LastCellNum; ++i)
        {
            table.Columns.Add(headerRow.GetCell(i).ToString());
        }

        // populate datatable rows with non-header rows
        for (int i = sheet.FirstRowNum + 1; i <= sheet.LastRowNum; ++i)
        {
            var row = sheet.GetRow(i);
            var tableRow = table.NewRow();
            for (int j = 0; j < row.Cells.Count; ++j)
            {
                var cell = row.GetCell(j);

                if (cell is not null)
                {
                    switch (cell.CellType)
                    {
                        case CellType.String:
                            tableRow[j] = cell.StringCellValue;
                            break;
                        case CellType.Blank:
                            tableRow[j] = string.Empty;
                            break;
                        case CellType.Numeric:
                            tableRow[j] = DateUtil.IsCellDateFormatted(cell) ? cell.DateCellValue : cell.NumericCellValue;
                            break;
                        case CellType.Boolean:
                            tableRow[j] = cell.BooleanCellValue;
                            break;
                    }
                }
            }
            table.Rows.Add(tableRow);
        }
        WriteToDb(table);
    }

    private void WriteToDb(DataTable table)
    {
        Console.WriteLine("File parsed successfully, writing to database");
        throw new NotImplementedException();
        //using var stream = new SqliteConnection($"Data Source={_name}");
    }
}
