using TableParser.Data;

namespace TableParser.Output
{
	public class EntriesToCsvWriter : BaseOutputWriter
	{
		public EntriesToCsvWriter(string outputPath, string fileName) : base(outputPath, fileName)
		{ 
		}

		public void Write(IEnumerable<EntryDescription> entries)
		{
			CreateDirectoryIfNeeded();

			using var writer = new StreamWriter(GetFilePath());

			writer.WriteLine("Description,Unit,Count");

			foreach (var entry in entries)
			{
				writer.WriteLine($"{entry.Name} {entry.Description},{entry.Unit},{entry.Count}");
			}
		}
	}
}


