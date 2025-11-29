using System.Text;
using TableParser.Data;

namespace TableParser.Output
{
	public class WorksheetsToCsvWriter : BaseOutputWriter, IOutputWriter
	{
		public WorksheetsToCsvWriter(string outputPath, string fileName) : base(outputPath, fileName)
		{
		}

		public void Write(IEnumerable<string> fileNames, IReadOnlyDictionary<EntryKey, Dictionary<string, double>> entries)
		{
			CreateDirectoryIfNeeded();

			var counts = new Dictionary<EntryKey, double>();
			foreach (var (key, countsByFile) in entries)
			{
				counts[key] = countsByFile.Values.Sum();
			}

			var orderedDescendingKeys = counts.Keys.OrderByDescending(key => counts[key]);

			using var writer = new StreamWriter(GetFilePath(), false, Encoding.UTF8);
			var stringBuilder = new StringBuilder();
			stringBuilder.Append("Объект,Габариты,Всего");

			foreach (var fileName in fileNames)
			{
				stringBuilder.Append($",{fileName}");
			}
			
			writer.WriteLine(stringBuilder.ToString());

			foreach (var key in orderedDescendingKeys)
			{
				stringBuilder.Clear();

				var countsByFile = entries[key];

				stringBuilder.Append($"{key.Name},{key.Description},{counts[key]}");
				foreach (var fileName in fileNames)
				{
					var countText = "-";
					if (countsByFile.TryGetValue(fileName, out var count))
					{
						countText = count.ToString();
					}
					stringBuilder.Append($",{countText}");
				}
				
				writer.WriteLine(stringBuilder.ToString());
			}
		}
	}
}


