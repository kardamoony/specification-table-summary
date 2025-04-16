using System;

namespace TableParser.Data
{
	public class SheetDescription
	{
		private Dictionary<string, double> _items;

		public string Name { get; }
		public IReadOnlyDictionary<string, double> Items => _items;

		public SheetDescription(string name)
		{
			Name = name;
			_items = new Dictionary<string, double>();
		}

		//public void 

	}
}


