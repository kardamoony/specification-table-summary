namespace TableParser
{
	public interface IValueParser<T>
	{
		bool TryParse(string value, out T result);
	}
}


