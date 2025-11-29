using Newtonsoft.Json;
using OfficeOpenXml;
using TableParser;
using TableParser.Data;

internal class Program
{
	private static string ConfigsPath = "configs";
	
	private static void Main(string[] args)
	{
		//required for OfficeOpenXml package to work
		ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
		
		var configsPath = Path.Combine(AppContext.BaseDirectory, ConfigsPath);

		if (!Directory.Exists(configsPath))
		{
			throw new FileNotFoundException("Configs directory doesn't exist: " + configsPath);
		}
		
		var files = Directory.GetFiles(configsPath);
		var parsingOptions = new List<Config>();
		
		foreach (var filePath in files)
		{
			var configFile = File.ReadAllText(filePath);

			try
			{
				var config = JsonConvert.DeserializeObject<Config>(configFile);
				var parserType = config.FactoryConfig.ParserType;
				if (!string.IsNullOrEmpty(parserType))
				{
					parsingOptions.Add(config);
				}
			}
			catch (Exception e)
			{
				Console.WriteLine(e);
			}
		}

		while (DoLoop(parsingOptions))
		{
			//loop
		}
	}

	private static bool DoLoop(IReadOnlyList<Config> options)
	{
		PrintOptions(options);
		var input = Console.ReadLine();

		if (!TryParseInput(input!, options, out var config, out var quitRequested))
		{
			Console.WriteLine($"Unknown command: {input}");
			return true;
		}
		
		if (quitRequested)
		{
			Console.WriteLine("Bye!");
			return false;
		}

		var hasParser = ComponentsFactory.TryGetParser(config, out var parser);
		var hasWriter = ComponentsFactory.TryGetWriter(config, out var writer, "summary-");

		if (hasParser && hasWriter)
		{
			var parsingLogic = new ParsingLogic(new ParsingLogic.Ctx
			{
				Config = config,
				Parser = parser!
			});
			
			var parsed = parsingLogic.Parse();
			if (parsed.Entries.Count > 0)
			{
				writer!.Write(parsed.Groups, parsed.Entries);
				Console.WriteLine("\n--- Done! ---\n");
			}
		} 
		else
		{
			Console.WriteLine($"Failed to create one of components. Has parser={hasParser}, has writer={hasWriter}");
		}

		return true;
	}
	
	private static void PrintOptions(IReadOnlyList<Config> options)
	{
		var i = 1;
		foreach (var config in options)
		{
			var name = config.FactoryConfig.ParserName;
			Console.WriteLine($"{i}: {name}");
			i++;
		}
		
		Console.WriteLine("\nQ: quit\n");
	}

	private static bool TryParseInput(string input, IReadOnlyList<Config> options, out Config result, out bool quitRequested)
	{
		var trimmed = input.Trim();
		if (IsQuitRequest(input))
		{
			result = default;
			quitRequested = true;
			return true;
		}
		
		if (!int.TryParse(trimmed, out var value) || value < 1 || value > options.Count)
		{
			result = default;
			quitRequested = false;
			return false;
		}

		result = options[value - 1];
		quitRequested = false;
		return true;
	}

	private static bool IsQuitRequest(string input)
	{
		var lower = input.ToLower();
		switch (lower)
		{
			case "q":
			case "quit":
			case "exit":
			{
				return true;
			}
		}

		return false;
	}
}