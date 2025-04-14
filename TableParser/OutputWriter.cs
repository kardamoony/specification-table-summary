namespace TableParser
{
	public class OutputWriter
	{
		private readonly string path;

		public OutputWriter(string outputPath)
		{
			path = Environment.CurrentDirectory + Path.DirectorySeparatorChar + outputPath;
		}

		public void Write(IEnumerable<EntryDescription> entries)
		{
			if (!Directory.Exists(path))
			{
				Directory.CreateDirectory(path);
			}

			using var writer = new StreamWriter(path + Path.DirectorySeparatorChar + GetFileName());

			writer.WriteLine("Description,Unit,Count");

			foreach (var entry in entries)
			{
				writer.WriteLine($"{entry.Name},{entry.Unit},{entry.Count}");
			}
		}

		private string GetFileName()
		{
			var now = DateTime.Now.ToString("yyyy-MM-dd-HH-mm-ss");
			return "summary-" + now + ".csv";
		}
	}
}


