namespace TableParser.Data
{
	public class EntryDescription
	{
		public string Name;
		public string Description;
		public string Unit;
		public double Count;
	}

	public struct EntryKey
	{
		private readonly string _name;
		private readonly string _description;

		public EntryKey(string name, string description)
		{
			_name = name;
			_description = description;
		}

		public override int GetHashCode()
		{
			return _name.GetHashCode() ^ _description.GetHashCode();
		}

		public override string ToString()
		{
			return $"{_name} {_description}";
		}
	}
}