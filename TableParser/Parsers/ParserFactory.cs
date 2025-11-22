using System.Reflection;
using TableParser.Data;

namespace TableParser.Parsers;

public static class ParserFactory
{
    private static Dictionary<string, Type> _parserTypes = new();

    static ParserFactory()
    {
        var assembly = Assembly.GetExecutingAssembly();
        var types = GetAllTypesThatImplementInterface<IWorksheetParser>(assembly);
        foreach (var type in types)
        {
            _parserTypes.Add(type.FullName!, type);
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
    
    private static IEnumerable<Type> GetAllTypesThatImplementInterface<TInterface>(Assembly assembly)
    {
        return assembly
            .GetTypes()
            .Where(type => typeof(TInterface).IsAssignableFrom(type) && type is { IsInterface: false, IsAbstract: false });
    }
}