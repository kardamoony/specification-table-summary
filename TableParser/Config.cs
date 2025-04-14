internal struct Settings
{
	public string dimensions_format;
	public string diameter_format;
	public string input_path_name;
	public string output_path_name;
}

internal struct Config
{
	public Dictionary<string, List<string>> include_keys;
	public List<string> exclude_keys;
	public List<string> units; 
}