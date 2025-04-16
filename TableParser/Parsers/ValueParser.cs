using System.Text.RegularExpressions;

namespace TableParser.Parsers
{
	public class ValueParser<T> : IValueParser<T>
	{
		private readonly IEnumerable<Regex> _regexes;
		private readonly Func<string, T> _parserFunc;

		public ValueParser(IEnumerable<string> patterns, Func<string, T> parserFunc)
		{
			_regexes = patterns.Select(p => new Regex(p));
			_parserFunc = parserFunc;
		}

		public bool TryParse(string value, out T result)
		{
			foreach (var regex in _regexes)
			{
				var match = regex.Match(value);
				if (match.Success)
				{
					try
					{
						result = _parserFunc(match.Groups[1].Value);
						return true;
					}
					catch (Exception)
					{
						Console.WriteLine($"Error parsing value: {value}");
					}
				}
			}

			result = default;
			return false;
		}
	}
	
	public class ValueParser<T1, T2> : IValueParser<(T1, T2)>
	{
		private readonly IEnumerable<Regex> _dimensionsRegexes;
		private readonly Func<string, T1> _parserFunc1;
		private readonly Func<string, T2> _parserFunc2;

		public ValueParser(IEnumerable<string> patterns, Func<string, T1> parserFunc1, Func<string, T2> parserFunc2)
		{
			_dimensionsRegexes = patterns.Select(p => new Regex(p));
			_parserFunc1 = parserFunc1;
			_parserFunc2 = parserFunc2;
		}

		public bool TryParse(string value, out (T1, T2) result)
		{
			foreach (var regex in _dimensionsRegexes)
			{
				var match = regex.Match(value);
				if (match.Success)
				{
					try
					{
						var t1 = _parserFunc1(match.Groups[1].Value);
						var t2 = _parserFunc2(match.Groups[2].Value);
						result = (t1, t2);
						return true;
					}
					catch (Exception)
					{
						Console.WriteLine($"Error parsing value: {value}");
					}
				}
			}

			result = default;
			return false;
		}
	}
}