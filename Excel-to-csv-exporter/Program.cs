using System;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using ExcelDataReader;

class Program
{
    static void Main(string[] args)
    {
        if (args.Length < 2)
        {
            Console.WriteLine("Usage: ExcelToCsv <input.xlsx> <output.csv> [sheetName|sheetIndex] [delimiter] [range]");
            Console.WriteLine("  range example: A1:D100  (optional, inclusive, Excel-style)");
            return;
        }

        string inputPath = args[0];
        string outputPath = args[1];
        string? sheetArg = args.Length >= 3 ? args[2] : null;
        string delimiter = args.Length >= 4 ? args[3] : ",";
        string? rangeArg = args.Length >= 5 ? args[4] : null;

        // Parse range if provided
        (int startRow, int endRow, int startCol, int endCol)? range = null;
        if (!string.IsNullOrEmpty(rangeArg))
        {
            range = ParseRange(rangeArg);
            if (range == null)
            {
                Console.WriteLine($"⚠️ Invalid range '{rangeArg}'. Falling back to full sheet.");
            }
        }

        // Enable encoding support
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        using var stream = File.Open(inputPath, FileMode.Open, FileAccess.Read);
        using var reader = ExcelReaderFactory.CreateReader(stream);

        using var writer = new StreamWriter(outputPath, false, Encoding.UTF8);

        int sheetIndex = 0;
        bool matchedSheet = false;

        do
        {
            bool useThisSheet = false;

            if (sheetArg == null)
                useThisSheet = (sheetIndex == 0); // default: first sheet
            else if (int.TryParse(sheetArg, out int idx))
                useThisSheet = (sheetIndex == idx);
            else if (reader.Name.Equals(sheetArg, StringComparison.OrdinalIgnoreCase))
                useThisSheet = true;

            if (useThisSheet)
            {
                matchedSheet = true;

                int rowCounter = 0;
                while (reader.Read())
                {
                    int colCount = reader.FieldCount;

                    // If range provided → skip outside rows
                    if (range != null && (rowCounter < range.Value.startRow || rowCounter > range.Value.endRow))
                    {
                        rowCounter++;
                        continue;
                    }

                    int actualStartCol = range?.startCol ?? 0;
                    int actualEndCol = range?.endCol ?? (colCount - 1);

                    var row = new string[(actualEndCol - actualStartCol + 1)];

                    for (int i = actualStartCol; i <= actualEndCol && i < colCount; i++)
                    {
                        row[i - actualStartCol] = reader.GetValue(i)?.ToString()?.Replace("\"", "\"\"") ?? "";
                    }

                    writer.WriteLine(string.Join(delimiter, row));
                    rowCounter++;
                }
            }

            sheetIndex++;

        } while (reader.NextResult());

        if (!matchedSheet && sheetArg != null)
        {
            Console.WriteLine($"⚠️ Sheet '{sheetArg}' not found. No data written.");
        }
    }

    // Parse Excel-like ranges (A1:D100)
    private static (int startRow, int endRow, int startCol, int endCol)? ParseRange(string range)
    {
        var match = Regex.Match(range, @"^([A-Z]+)(\d+):([A-Z]+)(\d+)$", RegexOptions.IgnoreCase);
        if (!match.Success) return null;

        int startCol = ColumnToIndex(match.Groups[1].Value);
        int startRow = int.Parse(match.Groups[2].Value) - 1; // Excel is 1-based, we use 0-based
        int endCol = ColumnToIndex(match.Groups[3].Value);
        int endRow = int.Parse(match.Groups[4].Value) - 1;

        return (startRow, endRow, startCol, endCol);
    }

    // Convert Excel column letters (A,B,C,...,AA) to zero-based index
    private static int ColumnToIndex(string colLetters)
    {
        int col = 0;
        foreach (char c in colLetters.ToUpper())
        {
            col *= 26;
            col += (c - 'A' + 1);
        }
        return col - 1;
    }
}
