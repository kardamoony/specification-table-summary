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
		public readonly string Name;
		public readonly string Description;

		public EntryKey(string name, string description)
		{
			Name = name;
			Description = description;
		}

		public override int GetHashCode()
		{
			return Name.GetHashCode() ^ Description.GetHashCode();
		}

		public override string ToString()
		{
			return $"{Name} {Description}";
		}
	}
}