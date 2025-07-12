using ExcelDateImporter.Api.Models;
using ClosedXML.Excel;
using System.Globalization;

namespace ExcelDateImporter.Api.Services
{
    public class ExcelDateParsingService
    {
        private readonly string[] _supportedFormats = {
            "dd/MM/yyyy",
            "MM/dd/yyyy", 
            "yyyy-MM-dd",
            "dd-MM-yyyy",
            "MM-dd-yyyy",
            "dd.MM.yyyy",
            "MM.dd.yyyy",
            "d/M/yyyy",
            "M/d/yyyy",
            "yyyy/MM/dd",
            "dd MMM yyyy",
            "MMM dd, yyyy"
        };

        public async Task<List<ExcelDateEntry>> ProcessExcelFile(IFormFile file)
        {
            var results = new List<ExcelDateEntry>();

            using var stream = new MemoryStream();
            await file.CopyToAsync(stream);
            
            using var workbook = new XLWorkbook(stream);
            var worksheet = workbook.Worksheet(1);
            
            var usedRange = worksheet.RangeUsed();
            if (usedRange == null)
                return results;

            for (int row = usedRange.FirstRow().RowNumber(); row <= usedRange.LastRow().RowNumber(); row++)
            {
                for (int col = usedRange.FirstColumn().ColumnNumber(); col <= usedRange.LastColumn().ColumnNumber(); col++)
                {
                    var cell = worksheet.Cell(row, col);
                    if (!cell.IsEmpty())
                    {
                        var entry = ParseDateValue(cell.Value);
                        results.Add(entry);
                    }
                }
            }

            return results;
        }

        public ExcelDateEntry ParseDateValue(object cellValue)
        {
            var entry = new ExcelDateEntry
            {
                OriginalValue = cellValue.ToString() ?? string.Empty
            };

            try
            {
                if (cellValue is DateTime dateTime)
                {
                    entry.ParsedDate = dateTime;
                    entry.DateFormat = "Excel DateTime";
                    entry.IsSuccessfullyParsed = true;
                    return entry;
                }

                if (cellValue is double doubleValue)
                {
                    if (IsExcelSerialDate(doubleValue))
                    {
                        entry.ParsedDate = DateTime.FromOADate(doubleValue);
                        entry.DateFormat = "Excel Serial Date";
                        entry.IsSuccessfullyParsed = true;
                        return entry;
                    }
                }

                var stringValue = cellValue.ToString();
                if (!string.IsNullOrWhiteSpace(stringValue))
                {
                    foreach (var format in _supportedFormats)
                    {
                        if (DateTime.TryParseExact(stringValue, format, CultureInfo.InvariantCulture, DateTimeStyles.None, out var parsedDate))
                        {
                            entry.ParsedDate = parsedDate;
                            entry.DateFormat = format;
                            entry.IsSuccessfullyParsed = true;
                            return entry;
                        }
                    }

                    if (DateTime.TryParse(stringValue, out var generalParsedDate))
                    {
                        entry.ParsedDate = generalParsedDate;
                        entry.DateFormat = "General Parse";
                        entry.IsSuccessfullyParsed = true;
                        return entry;
                    }
                }

                entry.ErrorMessage = "Unable to parse as date";
            }
            catch (Exception ex)
            {
                entry.ErrorMessage = ex.Message;
            }

            return entry;
        }

        private bool IsExcelSerialDate(double value)
        {
            return value >= 1 && value <= 2958465;
        }

        public string GetCSharpCodeExample()
        {
            return @"
// C# Code for Excel Date Processing using ClosedXML (Free & Open Source)
using ClosedXML.Excel;
using System.Globalization;

public class ExcelDateParser
{
    private readonly string[] _supportedFormats = {
        ""dd/MM/yyyy"", ""MM/dd/yyyy"", ""yyyy-MM-dd"",
        ""dd-MM-yyyy"", ""MM-dd-yyyy"", ""dd.MM.yyyy"",
        ""MM.dd.yyyy"", ""d/M/yyyy"", ""M/d/yyyy"",
        ""yyyy/MM/dd"", ""dd MMM yyyy"", ""MMM dd, yyyy""
    };

    public DateTime? ParseExcelDate(object cellValue)
    {
        try
        {
            // Handle Excel DateTime objects
            if (cellValue is DateTime dateTime)
                return dateTime;

            // Handle Excel serial dates (numeric values)
            if (cellValue is double doubleValue && IsExcelSerialDate(doubleValue))
                return DateTime.FromOADate(doubleValue);

            // Handle string dates with various formats
            var stringValue = cellValue.ToString();
            if (!string.IsNullOrWhiteSpace(stringValue))
            {
                // Try specific formats first
                foreach (var format in _supportedFormats)
                {
                    if (DateTime.TryParseExact(stringValue, format, 
                        CultureInfo.InvariantCulture, DateTimeStyles.None, 
                        out var parsedDate))
                        return parsedDate;
                }

                // Try general parsing as fallback
                if (DateTime.TryParse(stringValue, out var generalParsedDate))
                    return generalParsedDate;
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($""Error parsing date: {ex.Message}"");
        }

        return null;
    }

    private bool IsExcelSerialDate(double value)
    {
        // Excel serial dates range from 1 (1900-01-01) to 2958465 (9999-12-31)
        return value >= 1 && value <= 2958465;
    }
}

// Usage Example with ClosedXML:
var parser = new ExcelDateParser();

using var workbook = new XLWorkbook(""sample.xlsx"");
var worksheet = workbook.Worksheet(1);
var usedRange = worksheet.RangeUsed();

if (usedRange != null)
{
    for (int row = usedRange.FirstRow().RowNumber(); row <= usedRange.LastRow().RowNumber(); row++)
    {
        for (int col = usedRange.FirstColumn().ColumnNumber(); col <= usedRange.LastColumn().ColumnNumber(); col++)
        {
            var cell = worksheet.Cell(row, col);
            if (!cell.IsEmpty())
            {
                var parsedDate = parser.ParseExcelDate(cell.Value);
                
                if (parsedDate.HasValue)
                    Console.WriteLine($""Parsed: {parsedDate.Value:yyyy-MM-dd}"");
                else
                    Console.WriteLine($""Could not parse: {cell.Value}"");
            }
        }
    }
}";
        }
    }
}