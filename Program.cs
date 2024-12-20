﻿using System.Text;
using YamlDotNet.Serialization;
using YamlDotNet.Serialization.NamingConventions;
using ExcelMacro;
using Settings;
using Args;
using System.Reflection;

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
ComArgs comArgs = new ComArgs(args);
try
{
    comArgs.Validate();
}
catch (Exception e)
{
    Console.WriteLine($"エラー: {e.Message}");
    Console.WriteLine("--help オプションで引数の詳細をご覧ください。");
    return;
}

if (comArgs.Mode == "version")
{
    var version = Assembly.GetExecutingAssembly().GetName().Version;
    Console.WriteLine($"Version: {version}");
    return;
}

if (comArgs.Mode == "help")
{
    Console.WriteLine("exxm [target] mode\n");
    Console.WriteLine("target: 抽出または書き戻しの対象となる Excel ブック名。");
    Console.WriteLine("mode: 以下のいずれか。");
    Console.WriteLine("  --version: バージョン情報を表示します。");
    Console.WriteLine("  --help: このヘルプを表示します。");
    Console.WriteLine("  --from-xl: Excel ブックから VBA マクロを抽出します。");
    Console.WriteLine("  --to-xl: Excel ブックへ VBA マクロを書き戻します。");
    // TODO: --clean オプションを実装したらコメントアウトを外す
    //Console.WriteLine("  --clean: 抽出先または書き戻し先を初期化してから実行します。");
    return;
}

var macroIO = new MacroIO(settings, comArgs.Target);
var aborted = false;

if (macroIO.WbFiles.Count == 0)
{
    Console.WriteLine("対象となる Excel ブックがありません。");
    aborted = true;
}
else if (comArgs.Mode == "from-xl")
{
    /* Excel ブックから VBA マクロを抽出する */
    try
    {
        macroIO.ExtractMacros(comArgs.Clean);
    }
    catch (Exception e)
    {
        Console.WriteLine($"エラー: {e.Message}");
        aborted = true;
    }
}
else if (comArgs.Mode == "to-xl")
{
    /* Excel ブックへ VBA マクロを書き戻す */
    try
    {
        macroIO.WriteMacros(comArgs.Clean);
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
