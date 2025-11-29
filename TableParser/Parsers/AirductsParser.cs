using OfficeOpenXml;
using TableParser.Data;

namespace TableParser.Parsers;

public class AirductsParser : BaseParser
{
    private struct ParsedRowData
    {
        public string CellText;
        public double Area;
        public int Row;
        public bool NeedManualCheck;
    }
    
    private readonly IValueParser<double> _diameterParser;
    private readonly IValueParser<(double, double)> _dimensionsParser;

    private readonly HashSet<string> _units;

    public AirductsParser(Config config) : base(config)
    {
        _diameterParser = new ValueParser<double>(config.Settings.DiameterPatterns, double.Parse);
        _dimensionsParser = new ValueParser<double, double>(config.Settings.DimensionsPatterns, double.Parse, double.Parse);
    }
    
    public override void Parse(ExcelWorksheet worksheet, string fileName, ref Dictionary<EntryKey, EntryDescription> entries)
    {
        fileName = fileName.Replace(',', '_');
        
        var start = worksheet.Dimension.Start;
        var end = worksheet.Dimension.End;

        double cumulativeArea = 0;
        var needManualCheck = new List<int>();

        for (var row = start.Row; row <= end.Row; row++)
        {
            if (!TryProcessRow(worksheet, row, start, end, out var rowData))
            {
                continue;
            }

            if (rowData.NeedManualCheck)
            {
                needManualCheck.Add(rowData.Row);
                continue;
            }

            cumulativeArea += rowData.Area;
        }

        if (cumulativeArea > double.Epsilon)
        {
            var countedKey = new EntryKey(fileName, string.Empty);
            var entry = new EntryDescription
            {
                Count = cumulativeArea,
            };
            entries.Add(countedKey, entry);
        }

        foreach (var row in needManualCheck)
        {
            var message = $" [undefined]: row={row}";
            var key = new EntryKey(fileName, message);
            entries.Add(key, new EntryDescription());
        }
    }

    private bool TryProcessRow(ExcelWorksheet worksheet, int row, ExcelCellAddress start, ExcelCellAddress end, out ParsedRowData parsedData)
    {
        string? entryName = null;
        double? length = null;
        
        double? diameter = null;
        (double, double)? dimensions = null;
        
        var cellIsPossibleLength = false;
        
        for (var column = start.Column; column <= end.Column; column++)
        {
            var cell = worksheet.Cells[row, column];
            var cellText = cell.Text?.ToLowerInvariant().Trim();

            if (string.IsNullOrEmpty(cellText))
            {
                continue;
            }
            
            if (IsExcludeValue(cellText))
            {
                parsedData = default;
                return false;
            }

            if (cellIsPossibleLength && double.TryParse(cellText, out var l))
            {
                length = l;
                cellIsPossibleLength = false;
                continue;
            }

            if (IsIncludeValue(cellText, out var e))
            {
                entryName = e;

                if (_diameterParser.TryParse(cellText, out var d))
                {
                    diameter = d;
                    continue;
                }

                if (_dimensionsParser.TryParse(cellText, out var result))
                {
                    dimensions = result;
                    continue;
                }
            }

            if (IsUnit(cellText) && !length.HasValue)
            {
                cellIsPossibleLength = true;
            }
        }

        if (string.IsNullOrEmpty(entryName))
        {
            parsedData = default;
            return false;
        }

        if (length == null || (dimensions == null && diameter == null))
        {
            parsedData = new ParsedRowData
            {
                Row = row,
                NeedManualCheck = true,
            };
            return true;
        }

        if (diameter != null)
        {
            parsedData = new ParsedRowData
            {
                Area = GetArea(diameter.Value, length.Value),
                CellText = entryName,
                Row = row,
            };
            return true;
        }

        if (dimensions != null)
        {
            parsedData = new ParsedRowData
            {
                Area = GetArea(dimensions.Value.Item1, dimensions.Value.Item2, length.Value),
                CellText = entryName,
                Row = row,
            };
            return true;
        }

        parsedData = default;
        return false;
    }
 
    private static double GetArea(double diameter, double length)
    {
        return Math.PI * diameter * 0.001 * length;
    }

    private static double GetArea(double height, double width, double length)
    {
        return (height + width) * 0.001 * 2 * length;
    }
}