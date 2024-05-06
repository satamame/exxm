using System.Text;
using YamlDotNet.Serialization;
using YamlDotNet.Serialization.NamingConventions;
using ExcelMacro;
using Settings;

// エンコーディング プロバイダーを登録する
Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

/* 設定を読込む */
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
    Console.WriteLine("" +
        "エラー: --from-excel と --to-excel は同時に指定できません。");
    return;
}
if (!fromExcel && !toExcel)
{
    Console.WriteLine(
        "エラー: --from-excel または --to-excel を指定してください。");
    return;
}

var macroIO = new MacroIO(settings);
var aborted = false;

if (macroIO.WbFiles.Count == 0)
{
    Console.WriteLine("対象となる Excel ブックがありません。");
    aborted = true;
}
else if (fromExcel)
{
    /* Excel ブックから VBA マクロを抽出する */
    try
    {
        macroIO.ExtractMacros(clean);
    }
    catch (Exception e)
    {
        Console.WriteLine($"エラー: {e.Message}");
        aborted = true;
    }
}
else if (toExcel)
{
    /* Excel ブックへ VBA マクロを書き戻す */
    try
    {
        macroIO.WriteMacros(clean);
    }
    catch (Exception e)
    {
        Console.WriteLine($"エラー: {e.Message}");
        aborted = true;
    }
}

if (aborted)
{
    Console.WriteLine("処理が中断されました。");
}
else
{
    Console.WriteLine("処理が完了しました。");
}
