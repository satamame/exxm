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
    Console.WriteLine("エラー: --from-excel と --to-excel は同時に指定できません。");
    return;
}
if (!fromExcel && !toExcel)
{
    Console.WriteLine("エラー: --from-excel または --to-excel を指定してください。");
    return;
}

var macroIO = new MacroIO(settings);
var aborted = false;

if (fromExcel)
{
    /* Excel ブックから VBA マクロを抽出する */
    foreach (var f in macroIO.Files)
    {
        try
        {
            macroIO.ExtractMacros(f, clean);
        }
        catch (Exception e)
        {
            Console.WriteLine($"エラー: {e.Message}");
            aborted = true;
            break;
        }
    }
}
else if (toExcel)
{
    /* Excel ブックへ VBA マクロを書き戻す */
    foreach (var f in macroIO.Files)
    {
        try
        {
            macroIO.WriteBackMacros(f, clean);
        }
        catch (Exception e)
        {
            Console.WriteLine($"エラー: {e.Message}");
            aborted = true;
            break;
        }
    }
}

if (aborted)
{
    Console.WriteLine("処理が中断されました。");
    return;
}
Console.WriteLine("処理が完了しました。");
