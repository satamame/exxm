using YamlDotNet.Serialization;
using YamlDotNet.Serialization.NamingConventions;

Settings.Settings settings;
var deserializer = new DeserializerBuilder()
    .WithNamingConvention(CamelCaseNamingConvention.Instance)
    .Build();

using (var reader = new StreamReader("settings.yml"))
{
    settings = deserializer.Deserialize<Settings.Settings>(reader);
}

Console.WriteLine($"Excel Dir: {settings.Excel.Dir}");
Console.WriteLine($"Excel Recursive: {settings.Excel.Recursive}");
Console.WriteLine($"Excel Exclude: {string.Join(", ", settings.Excel.Exclude)}");
Console.WriteLine($"Excel Ext: {string.Join(", ", settings.Excel.Ext)}");
Console.WriteLine($"Macros Dir: {settings.Macros.Dir}");
