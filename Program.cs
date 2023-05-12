namespace SpreadsheetReader;

internal static class Program
{
    static void Main(string[] args)
    {
        if (args.Length == 0)
        {
            Console.WriteLine("Incorrect usage, please enter the filename as first argument");
            Console.WriteLine("Use -path followed by file path to select file not contained in current folder");
            Console.WriteLine("Use -dest followed by path to save to a different location");
            Console.ReadKey();
            return;
        }

        string path = "./";
        string dest = path;
        string filename = args[0];

        int i = 0;
        foreach (string arg in args)
        {
            if (arg.Contains("-path"))
            {
                path = args[++i];
                dest = path;
            }
            
            if (arg.Contains("-dest"))
            {
                dest = args[++i];

                if (!Directory.Exists(dest))
                {
                    Console.WriteLine("Destination not found");
                    return;
                }
            }

            i++;
        }

        if (!File.Exists(path + filename))
        {
            Console.WriteLine("File not found");
            return;
        }

        new SpreadsheetManager(path, filename, dest).ParseFile();
        Console.ReadKey();
    }

}