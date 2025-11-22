using OfficeOpenXml;
using TableParser.Data;
using TableParser.Parsers;

namespace TableParser
{
	public class ParsingLogic
	{
		private const string SheetExtension = ".xlsx";

		public struct Ctx
		{
			public Config Config;
			public IWorksheetParser Parser;
		}

		public struct Output
		{
			public IReadOnlyDictionary<EntryKey, Dictionary<string, double>> Entries;
			public IEnumerable<string> Groups;
		}

		private readonly Ctx _ctx;
		private readonly IWorksheetParser _worksheetParser;
		private readonly Dictionary<EntryKey, Dictionary<string, double>> _entries;
		private readonly HashSet<string> _groups;
		private readonly string _workingDirectory;

		public ParsingLogic(Ctx ctx)
		{
			_ctx = ctx;
			_entries = new ();
			_groups = new ();

			_worksheetParser = ctx.Parser;
			_workingDirectory = Path.Combine(AppContext.BaseDirectory, _ctx.Config.Settings.InputPath);
			if (!Directory.Exists(_workingDirectory))
			{
				Directory.CreateDirectory(_workingDirectory);
			}
		}

		public Output Parse()
		{
			_entries.Clear();
			_groups.Clear();

			foreach (var folder in Directory.GetDirectories(_workingDirectory))
			{
				var folderInfo = new DirectoryInfo(folder);
				var folderFiles = Directory.GetFiles(folder);
				_groups.Add(folderInfo.Name);
				if (folderFiles.Length == 0 || !folderFiles.Any(f => f.EndsWith(SheetExtension)))
				{
					continue;
				}

				foreach (var file in folderFiles)
				{
					var fileInfo = new FileInfo(file);
					if (fileInfo.Extension != SheetExtension) continue;
					ParseFile(fileInfo, folderInfo.Name);
				}
			}

			foreach (var file in Directory.GetFiles(_workingDirectory))
			{
				var fileInfo = new FileInfo(file);
				if (fileInfo.Extension != SheetExtension) continue;
				_groups.Add(fileInfo.Name);
				ParseFile(fileInfo, fileInfo.Name);
			}

			return new Output
			{
				Entries = new Dictionary<EntryKey, Dictionary<string, double>>(_entries),
				Groups = _groups.ToList(),
			};
		}

		private void ParseFile(FileInfo fileInfo, string groupName)
		{
			var fileEntries = new Dictionary<EntryKey, EntryDescription>();

			try
			{
				using var excelPackage = new ExcelPackage(fileInfo);
				var workbook = excelPackage.Workbook;

				foreach (var worksheet in workbook.Worksheets)
				{
					_worksheetParser.Parse(worksheet, fileInfo.Name, ref fileEntries);
				}
			}
			catch (Exception e)
			{
				Console.WriteLine(e);
			}

			foreach (var (key, entry) in fileEntries)
			{
				if (!_entries.TryGetValue(key, out var counts))
				{
					counts = new Dictionary<string, double>();
					_entries.Add(key, counts);
				}

				if (!counts.ContainsKey(groupName))
				{
					counts.Add(groupName, entry.Count);
				}
				else
				{
					counts[groupName] += entry.Count;
				}
			}
		}
	}
}


