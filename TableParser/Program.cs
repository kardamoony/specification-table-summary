using System.Text.RegularExpressions;
using Newtonsoft.Json;
using OfficeOpenXml;

internal class Program
{
	private static string CONFIG_PATH = Path.DirectorySeparatorChar + "config.json";
	private static string WORKING_DIRECTORY = Path.DirectorySeparatorChar + "sheets";
	private const string SHEET_EXTENSION = ".xlsx";

	internal class EntryDescription
	{
		public uint Id;
		public string Name;
		public string Unit;
		public double Count;
	}

	private const string DimensionsFormat = "{0}x{1}";
	private const string DiameterFormat = "ø{0}";

	private static readonly string[] _dimensionsPatterns = new [] 
	{
		@"(\d+)\s*[хХxX]\s*(\d+)", //cyrillic, latin
	};

	private static readonly string[] _diameterPatterns = new [] 
	{
		@"\u00f8\s*(\d+)",
		@"[Dd]\s*=\s*(\d+)",
		@"[Dd]\s*(\d+)",
		@"\d+",
	};

	private static Regex[] _dimensionsRegexes = _dimensionsPatterns.Select(p => new Regex(p)).ToArray();
	private static Regex[] _diameterRegexes = _diameterPatterns.Select(p => new Regex(p)).ToArray();

	private static Dictionary<string, HashSet<string>> _keys = new () //TODO: read from config file
	{
		{"Клапан противопожарный", new()
		{
			"клапан противопожарный", 
			"клапан противодымной вентиляции", 
			"огнезазадерживающий клапан",
			"клапан огнезадерживающий"
		}},
	};

	private static HashSet<string> _excludeKeys = new ()
	{
		"закрытый",
		"обогревом",
	};

	private static Dictionary<string, UnitType> _units = new () //TODO: read from config file
	{
		["шт"] = UnitType.Units,
		["кг"] = UnitType.Kilograms,
		["мп"] = UnitType.RunningMeters,
		["м²"] = UnitType.SquareMeters,
		["шт."] = UnitType.Units,
		["кг."] = UnitType.Kilograms,
		["м.п."] = UnitType.RunningMeters,
	};

	internal enum UnitType : byte
	{
		Unknown = 0,
		Units,
		Kilograms,
		RunningMeters,
		SquareMeters,
	}

	private static void Main(string[] args)
	{
		if (!File.Exists(Environment.CurrentDirectory + CONFIG_PATH))
		{
			throw new FileNotFoundException("Config file not found");
		}

		var configFile = File.ReadAllText(Environment.CurrentDirectory + CONFIG_PATH);
		var config = JsonConvert.DeserializeObject<Config>(configFile);

		ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
		var workingDirectory = Environment.CurrentDirectory + WORKING_DIRECTORY;
		if (!Directory.Exists(workingDirectory)) 
		{
			Directory.CreateDirectory(workingDirectory);
		}
		
		var files = Directory.GetFiles(workingDirectory);
		var entries = new Dictionary<string, EntryDescription>();

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
					var start = worksheet.Dimension.Start;
					var end = worksheet.Dimension.End;

					for (var row = start.Row; row <= end.Row; row++)
					{
						if (!TryExtractEntry(worksheet, row, out var entry))
						{
							continue;
						}

						if (!entries.TryGetValue(entry.Name, out var existingEntry))
						{
							entries.Add(entry.Name, entry);
							continue;
						}

						existingEntry.Count += entry.Count;
					}
				}
			}
			catch (Exception e)
			{
				//Console.WriteLine(e);
			}
			
		}

		foreach (var entry in entries)
		{
			Console.WriteLine($"{entry.Value.Name} [{entry.Value.Count} {entry.Value.Unit}]");
		}

		Console.WriteLine(workingDirectory);
		Console.ReadKey();
	}

	private static bool TryExtractEntry(ExcelWorksheet worksheet, int rowIdx, out EntryDescription entry)
	{
		var matchId = string.Empty;;
		var unit = string.Empty;
		var suffix = string.Empty;
		double count = 0;

		var isMatch = false;
		var extractCount = false;

		var startColumn = worksheet.Dimension.Start.Column;
		var endColumn = worksheet.Dimension.End.Column;

		for (var col = startColumn; col <= endColumn; col++)
		{
			var cell = worksheet.Cells[rowIdx, col];
			var cellText = cell.Text?.ToLowerInvariant().Trim();

			if (extractCount)
			{
				extractCount = false;
				count = (double)cell.Value;
				continue;
			}

			if (IsUnit(cellText, out var _))
			{
				unit = cellText;
				extractCount = true;

				continue;
			}

			if (!IsExcludeKey(cellText) && ContainsKey(cellText, out var id))
			{
				isMatch = true;
				matchId = id;

				if (TryExtractDimensions(cellText, out var w, out var h))
				{
					suffix = string.Format(DimensionsFormat, w, h);
				}
				else if (TryExtractDiameter(cellText, out var d))
				{
					suffix = string.Format(DiameterFormat, d);
				}
			}
		}

		if (isMatch)
		{
			entry = new EntryDescription
			{
				Name = matchId + " " + suffix,
				Unit = unit,
				Count = count,
			};
			return true;
		}

		entry = default;
		return false;
	}

	private static bool ContainsKey(string value, out string id)
	{
		id = string.Empty;
		if (string.IsNullOrEmpty(value)) return false;

		foreach (var pair in _keys)
		{
			foreach (var key in pair.Value)
			{
				if (value.Contains(key, StringComparison.InvariantCultureIgnoreCase)) 
				{
					id = pair.Key;
					return true;
				}
			}
		}
		return false;
	}

	private static bool IsUnit(string value, out UnitType unitType)
	{
		return _units.TryGetValue(value, out unitType);
	}

	private static bool IsExcludeKey(string value)
	{
		foreach (var key in _excludeKeys)
		{
			if (value.Contains(key, StringComparison.InvariantCultureIgnoreCase)) return true;
		}
		return false;
	}

	private static bool TryExtractDimensions(string value, out double width, out double height)
	{
		width = 0;
		height = 0;
		foreach (var regex in _dimensionsRegexes)
		{
			var match = regex.Match(value);
			if (match.Success)
			{
				width = double.Parse(match.Groups[1].Value);
				height = double.Parse(match.Groups[2].Value);
				return true;
			}
		}
		return false;
	}

	private static bool TryExtractDiameter(string value, out double diameter)
	{
		diameter = 0;
		foreach (var regex in _diameterRegexes)
		{
			var match = regex.Match(value);
			if (match.Success)
			{
				diameter = double.Parse(match.Groups[1].Value);
				return true;
			}
		}
		return false;
	}
}