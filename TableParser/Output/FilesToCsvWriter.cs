using System.Text;
using TableParser.Data;

namespace TableParser.Output
{
	public class FilesToCsvWriter : BaseOutputWriter
	{
		public FilesToCsvWriter(string outputPath, string fileName) : base(outputPath, fileName)
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

			using var writer = new StreamWriter(GetFilePath());
			var stringBuilder = new StringBuilder();
			stringBuilder.Append(",Всего");

			foreach (var fileName in fileNames)
			{
				stringBuilder.Append($",{fileName}");
			}
			
			writer.WriteLine(stringBuilder.ToString());

			foreach (var key in orderedDescendingKeys)
			{
				stringBuilder.Clear();

				var countsByFile = entries[key];

				stringBuilder.Append($"{key.ToString()},{counts[key]}");
				foreach (var fileName in fileNames)
				{
					countsByFile.TryGetValue(fileName, out var count);
					stringBuilder.Append($",{count}");
				}
				
				writer.WriteLine(stringBuilder.ToString());
			}
		}
	}
}


