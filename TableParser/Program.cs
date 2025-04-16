using Newtonsoft.Json;
using OfficeOpenXml;
using TableParser.Data;
using TableParser.Output;
using TableParser.Parsers;

//TODO: write error message if couldn't find UNIT
//TODO: support multiple files for 1 object

internal class Program
{
	private static string ConfigPath = "config.json";
	private const string SheetExtension = ".xlsx";

	private static void Main(string[] args)
	{
		if (!File.Exists(ConfigPath))
		{
			throw new FileNotFoundException("Config file not found: " + ConfigPath);
		}

		//required for OfficeOpenXml package to work
		ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;

		var configFile = File.ReadAllText(ConfigPath);
		var config = JsonConvert.DeserializeObject<Config>(configFile);

		var workingDirectory = config.Settings.InputPath;
		if (!Directory.Exists(workingDirectory))
		{
			Directory.CreateDirectory(workingDirectory);
			return;
		}

		var parser = new WorksheetParser(new WorksheetParser.Ctx
		{
			IncludeKeys = config.IncludeKeys,
			ExcludeKeys = config.ExcludeKeys,
			UnitsKeys = config.UnitsKeys,
			DimensionsParser = new ValueParser<double, double>(config.Settings.DimensionsPatterns, double.Parse, double.Parse),
			DiameterParser = new ValueParser<double>(config.Settings.DiameterPatterns, double.Parse),
			DimensionsFormat = config.Settings.DimensionsFormat,
			DiameterFormat = config.Settings.DiameterFormat,
		});

		var fileNames = new HashSet<string>();
		var entries = new Dictionary<EntryKey, Dictionary<string, double>>();

		foreach (var file in Directory.GetFiles(workingDirectory))
		{
			if (!file.EndsWith(SheetExtension)) continue;

			var fileName = Path.GetFileNameWithoutExtension(file);
			fileNames.Add(fileName);

			var fileEntries = new Dictionary<EntryKey, EntryDescription>();

			try
			{
				var fileInfo = new FileInfo(file);
				using var excelPackage = new ExcelPackage(fileInfo);
				var workbook = excelPackage.Workbook;

				foreach (var worksheet in workbook.Worksheets)
				{
					parser.Parse(worksheet, fileName, ref fileEntries);
				}
			}
			catch (Exception e)
			{
				Console.WriteLine(e);
			}

			foreach (var (key, entry) in fileEntries)
			{
				if (!entries.TryGetValue(key, out var counts))
				{
					counts = new Dictionary<string, double>();
					entries.Add(key, counts);
				}

				counts.Add(fileName, entry.Count);
			}
		}

		var writer = new FilesToCsvWriter(config.Settings.OutputPath, "summary-");
		writer.Write(fileNames, entries);

		Console.WriteLine("Done!");
		//Console.ReadKey();
	}
}