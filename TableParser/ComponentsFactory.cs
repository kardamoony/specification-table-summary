using System.Reflection;
using TableParser.Data;
using TableParser.Output;
using TableParser.Parsers;

namespace TableParser;

public static class ComponentsFactory
{
    private static Dictionary<string, Type> _parserTypes = new();
    private static Dictionary<string, Type> _outputWriterTypes = new();

    static ComponentsFactory()
    {
        var assembly = Assembly.GetExecutingAssembly();
        var parserTypes = GetAllTypesThatImplementInterface<IWorksheetParser>(assembly);
        var outputWriterTypes = GetAllTypesThatImplementInterface<IOutputWriter>(assembly);
        
        foreach (var type in parserTypes)
        {
            _parserTypes.Add(type.FullName!, type);
        }
        
        foreach (var type in outputWriterTypes)
        {
            _outputWriterTypes.Add(type.FullName!, type);
        }
    }
    
    public static bool TryGetParser(Config config, out IWorksheetParser? parser)
    {
        if (!_parserTypes.TryGetValue(config.FactoryConfig.ParserType, out var parserType))
        {
            parser = default;
            return false;
        }

        parser = Activator.CreateInstance(parserType, config) as IWorksheetParser;
        return true;
    }

    public static bool TryGetWriter(Config config, out IOutputWriter? outputWriter, string outputFileName)
    {
        if (!_outputWriterTypes.TryGetValue(config.FactoryConfig.OutputType, out var outputType))
        {
            outputWriter = default;
            return false;
        }

        var outputPath = Path.Combine(AppContext.BaseDirectory, config.Settings.OutputPath);
        outputWriter = Activator.CreateInstance(outputType, outputPath, outputFileName) as IOutputWriter;
        return true;
    }
    
    private static IEnumerable<Type> GetAllTypesThatImplementInterface<TInterface>(Assembly assembly)
    {
        return assembly
            .GetTypes()
            .Where(type => typeof(TInterface).IsAssignableFrom(type) && type is { IsInterface: false, IsAbstract: false });
    }
}