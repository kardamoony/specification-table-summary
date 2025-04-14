using System.Text.RegularExpressions;

namespace TableParser.ValueParsers
{
	public class DimensionsParser : IValueParser<(double, double)>
	{
		private readonly IEnumerable<string> _dimensionsPatterns = new [] 
		{
			@"(\d+)\s*[хХxX]\s*(\d+)", //cyrillic, latin
		};

		private readonly IEnumerable<Regex> _dimensionsRegexes;

		public DimensionsParser()
		{
			_dimensionsRegexes = _dimensionsPatterns.Select(p => new Regex(p));
		}

		public bool TryParse(string value, out (double, double) result)
		{
			foreach (var regex in _dimensionsRegexes)
			{
				var match = regex.Match(value);
				if (match.Success)
				{
					var width = double.Parse(match.Groups[1].Value);
					var height = double.Parse(match.Groups[2].Value);
					result = (width, height);
					return true;
				}
			}

			result = default;
			return false;
		}
	}
}


