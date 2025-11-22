using Newtonsoft.Json;
using OfficeOpenXml;
using TableParser;
using TableParser.Data;
using TableParser.Output;

internal class Program
{
	private static string ConfigPath = "config.json";
	
	private static void Main(string[] args)
	{
		var configPath = Path.Combine(AppContext.BaseDirectory, ConfigPath);
		if (!File.Exists(configPath))
		{
			throw new FileNotFoundException("Config file not found: " + configPath);
		}

		//required for OfficeOpenXml package to work
		ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;

		var configFile = File.ReadAllText(ConfigPath);
		var config = JsonConvert.DeserializeObject<Config>(configFile);

		var parsingLogic = new ParsingLogic(new ParsingLogic.Ctx{Config = config});
		var parsed = parsingLogic.Parse();

		if (parsed.Entries.Count < 1) return;

		var outputPath = Path.Combine(AppContext.BaseDirectory, config.Settings.OutputPath);
		var writer = new FilesToCsvWriter(outputPath, "summary-");
		writer.Write(parsed.Groups, parsed.Entries);

		Console.WriteLine("Done!");
		Console.ReadKey();
	}
}