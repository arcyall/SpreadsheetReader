namespace SpreadsheetReader;

internal static class Program
{
    static void Main(string[] args)
    {
        string? input;

        while (true)
        {
            Console.WriteLine("Enter filename with file extension: ");
            input = Console.ReadLine();

            if (File.Exists($"./{input}"))
            {
                Console.WriteLine("File found");
                break;
            }
            Console.WriteLine("File not found");
        }

        SpreadsheetManager.ParseFile(input);
        Console.ReadKey();
    }

}