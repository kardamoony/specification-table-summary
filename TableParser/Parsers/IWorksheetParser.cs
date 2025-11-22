using OfficeOpenXml;
using TableParser.Data;

namespace TableParser.Parsers;

public interface IWorksheetParser
{
    void Parse(ExcelWorksheet worksheet, string fileName, ref Dictionary<EntryKey, EntryDescription> entries);
}