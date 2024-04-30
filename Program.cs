using System.Text;
using YamlDotNet.Serialization;
using YamlDotNet.Serialization.NamingConventions;
using ExcelMacro;
using Settings;

// エンコーディング プロバイダーを登録する
Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

/* 設定ファイルを読込む */
AppSettings settings;
var deserializer = new DeserializerBuilder()
    .WithNamingConvention(CamelCaseNamingConvention.Instance)
    .Build();

using (var reader = new StreamReader("exxm-conf.yml"))
{
    settings = deserializer.Deserialize<AppSettings>(reader);
}

/* コマンドライン引数を取得する */
bool fromExcel = args.Contains("--from-excel");
bool toExcel = args.Contains("--to-excel");
bool clean = args.Contains("--clean");

if (fromExcel && toExcel)
{
    Console.WriteLine("エラー: --from-excel と --to-excel は同時に指定できません。");
    return;
}
if (!fromExcel && !toExcel)
{
    Console.WriteLine("エラー: --from-excel または --to-excel を指定してください。");
    return;
}

/* 対象となる Excel ファイルを取得する */
var files = ExcelMacroIO.FindExcelFiles(
    settings.Excel.Dir, settings.Excel.Exclude, settings.Excel.Ext);

if (fromExcel)
{
    /* Excel ブックから VBA マクロを抽出する */
    foreach (var f in files)
    {
        try
        {
            ExcelMacroIO.ExtractMacros(f, settings.Macros.Dir, clean);
        }
        catch (Exception e)
        {
            Console.WriteLine($"エラー: {e.Message}");
            break;
        }
    }
}
else if (toExcel)
{
    /* Excel ブックへ VBA マクロを書き戻す */
    foreach (var f in files)
    {
        try
        {
            ExcelMacroIO.WriteBackMacros(f, settings.Macros.Dir, clean);
        }
        catch (Exception e)
        {
            Console.WriteLine($"エラー: {e.Message}");
            break;
        }
    }
}
Console.WriteLine("処理が完了しました。");
