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

/* コマンドライン引数を取得する */
bool fromExcel = args.Contains("--from-excel");
bool toExcel = args.Contains("--to-excel");
bool clean = args.Contains("--clean");

if (fromExcel && toExcel)
{
    Console.WriteLine("エラー: --from-excel と --to-excel は同時に指定できません。");
    return;
}

var files = ExcelMacroIO.FindExcelFiles(
    settings.Excel.Dir, settings.Excel.Recursive, settings.Excel.Exclude,
    settings.Excel.Ext
);

if (fromExcel)
{
    foreach (var f in files)
    {
        try
        {
            ExcelMacroIO.ExtractMacros(f, clean);
        }
        catch (Exception e)
        {
            Console.WriteLine($"エラー: {e.Message}");
            break;
        }
    }
}

Console.WriteLine("何かキーを押して続行してください...");
Console.ReadLine();
