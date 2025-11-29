using OfficeOpenXml;
using TableParser.Data;

namespace TableParser.Parsers;

public abstract class BaseParser : IWorksheetParser
{
    private readonly IReadOnlyDictionary<string, List<string>> _includeKeys;
    private readonly IEnumerable<string> _excludeKeys;
    private readonly HashSet<string> _units;

    protected BaseParser(Config config)
    {
        _units = new HashSet<string>(config.UnitsKeys);
        _includeKeys = config.IncludeKeys;
        _excludeKeys = config.ExcludeKeys;
    }
    
    public abstract void Parse(ExcelWorksheet worksheet, string fileName, ref Dictionary<EntryKey, EntryDescription> entries);
    
    protected bool IsIncludeValue(string cellText, out string value)
    {
        value = string.Empty;
        if (string.IsNullOrEmpty(cellText) || _includeKeys == null) return false;

        var split = cellText.Split(' ', ',');

        foreach (var pair in _includeKeys)
        {
            foreach (var key in pair.Value)
            {
                foreach (var token in split)
                {
                    if (token.Equals(key, StringComparison.InvariantCultureIgnoreCase))
                    {
                        value = token;
                        return true;
                    }
                }
            }
        }
        
        return false;
    }

    protected bool IsExcludeValue(string value)
    {
        if (string.IsNullOrEmpty(value) || _excludeKeys == null) return false;
        var split = value.Split(' ', ',');

        foreach (var key in _excludeKeys)
        {
            foreach (var token in split)
            {
                if (token.Equals(key, StringComparison.InvariantCultureIgnoreCase))
                {
                    return true;
                }
            }
        }
        return false;
    }

    protected bool IsUnit(string value)
    {
        return !string.IsNullOrEmpty(value) && _units.Contains(value);
    }
}