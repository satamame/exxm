using System;
using YamlDotNet.Serialization;
using YamlDotNet.Serialization.NamingConventions;
using Settings;
using ExcelMacro;

/* 設定ファイルを読込む */
AppSettings settings;
var deserializer = new DeserializerBuilder()
    .WithNamingConvention(CamelCaseNamingConvention.Instance)
    .Build();

using (var reader = new StreamReader("settings.yml"))
{
    settings = deserializer.Deserialize<AppSettings>(reader);
}

Console.WriteLine($"Excel Dir: {settings.Excel.Dir}");
Console.WriteLine($"Excel Recursive: {settings.Excel.Recursive}");
Console.WriteLine($"Excel Exclude: {string.Join(", ", settings.Excel.Exclude)}");
Console.WriteLine($"Excel Ext: {string.Join(", ", settings.Excel.Ext)}");
Console.WriteLine($"Macros Dir: {settings.Macros.Dir}");

var files = ExcelMacroIO.FindExcelFiles(
    settings.Excel.Dir, settings.Excel.Recursive, settings.Excel.Exclude,
    settings.Excel.Ext
);
foreach (var f in files)
{
    // Open メソッドに絶対パスを渡さないとエラーになるため (原因不明)
    //string currentDir = Directory.GetCurrentDirectory();
    //string path = Path.Combine(currentDir, f);
    ExcelMacroIO.ExtractMacros(f);
}

Console.WriteLine("何かキーを押して続行してください...");
Console.ReadLine();
