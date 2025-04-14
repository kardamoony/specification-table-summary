using Newtonsoft.Json;
using OfficeOpenXml;
using TableParser;
using TableParser.ValueParsers;

internal class Program
{
	private static string CONFIG_PATH = Path.DirectorySeparatorChar + "config.json";
	private const string SHEET_EXTENSION = ".xlsx";


	private static void Main(string[] args)
	{
		if (!File.Exists(Environment.CurrentDirectory + CONFIG_PATH))
		{
			throw new FileNotFoundException("Config file not found");
		}

		var configFile = File.ReadAllText(Environment.CurrentDirectory + Path.DirectorySeparatorChar + CONFIG_PATH);
		var config = JsonConvert.DeserializeObject<Config>(configFile);

		ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
		var workingDirectory = Environment.CurrentDirectory + Path.DirectorySeparatorChar + config.Settings.InputPath;
		if (!Directory.Exists(workingDirectory)) 
		{
			Directory.CreateDirectory(workingDirectory);
		}
		
		var files = Directory.GetFiles(workingDirectory);
		var entries = new Dictionary<string, EntryDescription>();

		var parser = new WorksheetParser(new WorksheetParser.Ctx
		{
			IncludeKeys = config.IncludeKeys,
			ExcludeKeys = config.ExcludeKeys,
			UnitsKeys = config.UnitsKeys,
			DimensionsParser = new DimensionsParser(),
			DiameterParser = new DiameterParser(),
			DimensionsFormat = config.Settings.DimensionsFormat,
			DiameterFormat = config.Settings.DiameterFormat,
		});

		foreach (var file in files)
		{
			if (!file.EndsWith(SHEET_EXTENSION)) continue;

			try
			{
				var fileInfo = new FileInfo(file);
				using var excelPackage = new ExcelPackage(fileInfo);
				var workbook = excelPackage.Workbook;

				foreach (var worksheet in workbook.Worksheets)
				{
					parser.Parse(worksheet, ref entries);
				}
			}
			catch (Exception e)
			{
				Console.WriteLine(e);
			}
			
		}

		var writer = new OutputWriter(config.Settings.OutputPath);
		writer.Write(entries.Values);

		//Console.WriteLine(workingDirectory);
		//Console.ReadKey();
	}
}