using OfficeOpenXml;
using TableParser.Data;

namespace TableParser.Parsers
{
	public class WorksheetParser
	{
		private const string DefaultDescription = "[Undefined]";
		public struct Ctx
		{
			public IReadOnlyDictionary<string, List<string>> IncludeKeys;
			public IEnumerable<string> ExcludeKeys;
			public IEnumerable<string> UnitsKeys;
			public IValueParser<double> DiameterParser;
			public IValueParser<(double, double)> DimensionsParser;
			public string DimensionsFormat;
			public string DiameterFormat;
		}

		private readonly Ctx _ctx;
		private HashSet<string> _units;

		public WorksheetParser(Ctx ctx)
		{
			_ctx = ctx;
			_units = new (ctx.UnitsKeys);
		}

		public void Parse(ExcelWorksheet worksheet, string fileName, ref Dictionary<EntryKey, EntryDescription> entries)
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

				if (IsExcludeValue(cellText))
				{
					continue;
				}

				if (IsIncludeValue(cellText, out var id))
				{
					nameCellText = cellText;
					isMatch = true;
					name = id;

					if (_ctx.DimensionsParser.TryParse(cellText, out var data))
					{
						description = string.Format(_ctx.DimensionsFormat, data.Item1, data.Item2);
					}
					else if (_ctx.DiameterParser.TryParse(cellText, out var d))
					{
						description = string.Format(_ctx.DiameterFormat, d);
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

		private bool IsIncludeValue(string value, out string id)
		{
			id = string.Empty;
			if (string.IsNullOrEmpty(value)) return false;

			foreach (var pair in _ctx.IncludeKeys)
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

		private bool IsExcludeValue(string value)
		{
			if (string.IsNullOrEmpty(value)) return false;
			foreach (var key in _ctx.ExcludeKeys)
			{
				if (value.Contains(key, StringComparison.InvariantCultureIgnoreCase)) return true;
			}
			return false;
		}

		private bool IsUnit(string value)
		{
			return !string.IsNullOrEmpty(value) && _units.Contains(value);
		}
	}
}


