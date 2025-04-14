using Newtonsoft.Json;

public struct Settings
{
	[JsonProperty("dimensions_format")] public string DimensionsFormat;
	[JsonProperty("diameter_format")] public string DiameterFormat;
	[JsonProperty("input_path")] public string InputPath;
	[JsonProperty("output_path")] public string OutputPath;
}

public struct Config
{
	[JsonProperty("include_keys")] public Dictionary<string, List<string>> IncludeKeys;
	[JsonProperty("exclude_keys")] public List<string> ExcludeKeys;
	[JsonProperty("units_keys")] public List<string> UnitsKeys;
	[JsonProperty("settings")] public Settings Settings; 
}