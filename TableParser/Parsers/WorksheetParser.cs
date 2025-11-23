using OfficeOpenXml;
using TableParser.Data;

namespace TableParser.Parsers
{
	//used implicitly by Activator
	public class WorksheetParser : BaseParser
	{
		private const string DefaultDescription = "[Undefined]";
		
		private readonly IValueParser<double> _diameterParser;
		private readonly IValueParser<(double, double)> _dimensionsParser;
		private readonly string _dimensionsFormat;
		private readonly string _diameterFormat;
		
		public WorksheetParser(Config config) : base(config)
		{
			_diameterParser = new ValueParser<double>(config.Settings.DiameterPatterns, double.Parse);
			_dimensionsParser = new ValueParser<double, double>(config.Settings.DimensionsPatterns, double.Parse, double.Parse);
			_diameterFormat = config.Settings.DiameterFormat;
			_dimensionsFormat = config.Settings.DimensionsFormat;
		}

		public override void Parse(ExcelWorksheet worksheet, string fileName, ref Dictionary<EntryKey, EntryDescription> entries)
		{
			var start = worksheet.Dimension.Start;
			var end = worksheet.Dimension.End;

			for (var row = start.Row; row <= end.Row; row++)
			{
				if (!TryExtractEntry(worksheet, fileName, row, out var entry))
				{
					continue;
				}

				var key = new EntryKey(entry.Name, entry.Description);

				if (!entries.TryGetValue(key, out var existingEntry))
				{
					entries.Add(key, entry);
					continue;
				}

				existingEntry.Count += entry.Count;
			}
		}

		private bool TryExtractEntry(ExcelWorksheet worksheet, string fileName, int rowIdx, out EntryDescription entry)
		{
			var nameCellText = string.Empty;
			var name = string.Empty;
			var unit = string.Empty;
			var description = DefaultDescription;
			double count = 0;

			var isMatch = false;
			var extractCount = false;

			var startColumn = worksheet.Dimension.Start.Column;
			var endColumn = worksheet.Dimension.End.Column;

			for (var col = startColumn; col <= endColumn; col++)
			{
				var cell = worksheet.Cells[rowIdx, col];
				var cellText = cell.Text?.ToLowerInvariant().Trim();

				if (IsExcludeValue(cellText))
				{
					entry = default;
					return false;
				}

				if (extractCount)
				{
					extractCount = false;
					count = (double)cell.Value;
					continue;
				}

				if (IsUnit(cellText))
				{
					unit = cellText;
					extractCount = true;
					continue;
				}

				if (isMatch && description.Equals(DefaultDescription))
				{
					if (TryParseDescription(cellText, out var descr))
					{
						description = descr;
					}
				}

				if (IsIncludeValue(cellText, out var id))
				{
					nameCellText = cellText;
					isMatch = true;
					name = id;

					if (TryParseDescription(cellText, out var descr))
					{
						description = descr;
					}
				}
			}

			if (isMatch)
			{
				if (description.Equals(DefaultDescription))
				{
					Console.WriteLine($"Failed to parse description for [{nameCellText}] at {fileName}, row {rowIdx}");
				}

				if (count < double.Epsilon)
				{
					Console.WriteLine($"Failed to parse count for [{nameCellText}] at {fileName}, row {rowIdx}");
				}

				entry = new EntryDescription
				{
					Name = name,
					Description = description,
					Unit = unit,
					Count = count,
				};
				return true;
			}

			entry = default;
			return false;
		}

		private bool TryParseDescription(string value, out string description)
		{
			if (_dimensionsParser.TryParse(value, out var data))
			{
				description = string.Format(_dimensionsFormat, data.Item1, data.Item2);
				return true;
			}

			if (_diameterParser.TryParse(value, out var d))
			{
				description = string.Format(_diameterFormat, d);
				return true;
			}

			description = string.Empty;
			return false;
		}

		
	}
}


