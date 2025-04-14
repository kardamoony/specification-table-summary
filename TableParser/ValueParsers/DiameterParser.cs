using System.Text.RegularExpressions;

namespace TableParser.ValueParsers
{
	public class DiameterParser : IValueParser<double>
	{
		private readonly IEnumerable<string> _diameterPatterns = new [] 
		{
			@"\u00f8\s*(\d+)",
			@"[Dd]\s*=\s*(\d+)",
			@"[Dd]\s*(\d+)",
			@"\s+(\d+)\s+",
		};

		private readonly IEnumerable<Regex> _diameterRegexes;

		public DiameterParser()
		{
			_diameterRegexes = _diameterPatterns.Select(p => new Regex(p));
		}

		public bool TryParse(string value, out double result)
		{
			foreach (var regex in _diameterRegexes)
			{
				var match = regex.Match(value);
				if (match.Success)
				{
					result = double.Parse(match.Groups[1].Value);
					return true;
				}
			}

			result = 0;
			return false;
		}
	}
}
