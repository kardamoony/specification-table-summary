using Newtonsoft.Json;

namespace TableParser.Data
{
	public struct Settings
	{
		[JsonProperty("dimensions_format")] public string DimensionsFormat;
		[JsonProperty("dimensions_patterns")] public List<string> DimensionsPatterns;
		[JsonProperty("diameter_patterns")] public List<string> DiameterPatterns;
		[JsonProperty("diameter_format")] public string DiameterFormat;
		[JsonProperty("input_path")] public string InputPath;
		[JsonProperty("output_path")] public string OutputPath;
	}

	public struct FactoryConfig
	{
		[JsonProperty("parser_type")] public string ParserType;
		[JsonProperty("parser_name")] public string ParserName;
	}

	public struct Config
	{
		[JsonProperty("include_keys")] public Dictionary<string, List<string>> IncludeKeys;
		[JsonProperty("exclude_keys")] public List<string> ExcludeKeys;
		[JsonProperty("units_keys")] public List<string> UnitsKeys;
		[JsonProperty("settings")] public Settings Settings;
		[JsonProperty("factory_config")] public FactoryConfig FactoryConfig;
	}
}

