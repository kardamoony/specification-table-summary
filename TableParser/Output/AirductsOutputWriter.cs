using System.Text;
using TableParser.Data;

namespace TableParser.Output;

public class AirductsOutputWriter : BaseOutputWriter, IOutputWriter
{
    public AirductsOutputWriter(string outputPath, string fileName) : base(outputPath, fileName)
    {
    }
    
    public void Write(IEnumerable<string> fileNames, IReadOnlyDictionary<EntryKey, Dictionary<string, double>> entries)
    {
        CreateDirectoryIfNeeded();

        var ordered = entries.OrderBy(k => k.Key.Name);
        
        using var writer = new StreamWriter(GetFilePath(), false, Encoding.UTF8);
        var stringBuilder = new StringBuilder();
        stringBuilder.Append("Объект,Всего");
        writer.WriteLine(stringBuilder.ToString());

        foreach (var (key, data) in ordered)
        {
            if (key.Description.Contains("[undefined]"))
            {
                stringBuilder.Clear();
                stringBuilder.Append($"{key.Name},{key.Description}");
                writer.WriteLine(stringBuilder.ToString());
                continue;
            }
            
            foreach (var (_, count) in data)
            {
                stringBuilder.Clear();
                stringBuilder.Append($"{key.Name},{count.ToString("0.0")}");
                writer.WriteLine(stringBuilder.ToString());
            }
        }
    }
}