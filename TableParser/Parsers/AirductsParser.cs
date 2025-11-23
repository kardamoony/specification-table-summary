using OfficeOpenXml;
using TableParser.Data;

namespace TableParser.Parsers;

public class AirductsParser : BaseParser
{
    private const string UnitName = "м²";

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
        var start = worksheet.Dimension.Start;
        var end = worksheet.Dimension.End;

        double cumulativeArea = 0;
        var needManualCheck = new List<(string data, int row)>();

        for (var row = start.Row; row <= end.Row; row++)
        {
            if (!TryProcessRow(worksheet, row, start, end, out var rowData))
            {
                continue;
            }

            if (rowData.NeedManualCheck)
            {
                needManualCheck.Add((rowData.CellText, rowData.Row));
                continue;
            }

            cumulativeArea += rowData.Area;
        }
    }

    private bool TryProcessRow(ExcelWorksheet worksheet, int row, ExcelCellAddress start, ExcelCellAddress end, out ParsedRowData parsedData)
    {
        string entryName = null;
        double? length = null;
        
        double? diameter = null;
        (double, double)? dimensions = null;
        
        var cellIsPossibleLength = false;
        
        for (var column = start.Column; column <= end.Column; column++)
        {
            var cell = worksheet.Cells[row, column];
            var cellText = cell.Text?.ToLowerInvariant().Trim();
            
            if (cellText == null || IsExcludeValue(cellText))
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

        if (entryName == null)
        {
            parsedData = default;
            return false;
        }

        if (length == null || (dimensions == null && diameter == null))
        {
            parsedData = new ParsedRowData
            {
                CellText = entryName,
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
        return Math.PI * diameter * length;
    }

    private static double GetArea(double height, double width, double length)
    {
        return (height + width) * 2 * length;
    }
}