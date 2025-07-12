using ExcelDateImporter.Api.Models;
using OfficeOpenXml;
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
            
            ExcelPackage.License = LicenseContext.NonCommercial;
            
            using var package = new ExcelPackage(stream);
            var worksheet = package.Workbook.Worksheets[0];
            
            var rowCount = worksheet.Dimension?.Rows ?? 0;
            var colCount = worksheet.Dimension?.Columns ?? 0;

            for (int row = 1; row <= rowCount; row++)
            {
                for (int col = 1; col <= colCount; col++)
                {
                    var cellValue = worksheet.Cells[row, col].Value;
                    if (cellValue != null)
                    {
                        var entry = ParseDateValue(cellValue);
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
// C# Code for Excel Date Processing
using OfficeOpenXml;
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

// Usage Example:
var parser = new ExcelDateParser();
var excelFile = new FileInfo(""sample.xlsx"");

using var package = new ExcelPackage(excelFile);
var worksheet = package.Workbook.Worksheets[0];

for (int row = 1; row <= worksheet.Dimension.Rows; row++)
{
    for (int col = 1; col <= worksheet.Dimension.Columns; col++)
    {
        var cellValue = worksheet.Cells[row, col].Value;
        var parsedDate = parser.ParseExcelDate(cellValue);
        
        if (parsedDate.HasValue)
            Console.WriteLine($""Parsed: {parsedDate.Value:yyyy-MM-dd}"");
        else
            Console.WriteLine($""Could not parse: {cellValue}"");
    }
}";
        }
    }
}