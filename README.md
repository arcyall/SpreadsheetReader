# Readme

This repository contains code for a C# console application that reads an Excel spreadsheet file and loads it into an SQLite database. 

## Build

1. Clone the repository to your local machine. 
2. Open the solution file (`SpreadsheetReader.sln`) in Visual Studio. 
3. Build the project. 

## Usage

To run the program, use the following command:

```
SpreadsheetReader.exe <filename> [-path <path>] [-dest <destination>]
```

where `<filename>` is the name of the Excel spreadsheet file to be read.

### Optional Arguments

- `-path`: Specifies the path to the file to be read. By default, the program looks for the file in the current directory.
- `-dest`: Specifies the destination directory to which the database file should be saved. By default, the database file is saved in the same directory as the input file.

### Example

To read the file `employees.xlsx` located in the folder `C:\Users\Me\Documents`, and save the resulting database file in the same folder, use the following command:

```
SpreadsheetReader.exe employees.xlsx -path C:\Users\Me\Documents -dest C:\Users\Me\Documents
```

## Contributing

Pull requests are welcome. For major changes, please open an issue first to discuss what you would like to change.

## License

[MIT](https://choosealicense.com/licenses/mit/)