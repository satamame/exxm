namespace Settings;

// TODO: バリデーションのメソッドを実装する

public class AppSettings
{
    public ExcelSettings Excel { get; set; } = new ExcelSettings();
    public MacrosSettings Macros { get; set; } = new MacrosSettings();
}

public class ExcelSettings
{
    public string Dir { get; set; } = "books";
    public List<string> Exclude { get; set; } = [];
    public List<string> Ext { get; set; } = [".xlsm", ".xlsb"];
}

public class MacrosSettings
{
    public string Dir { get; set; } = "macros";
    // ブックごとのディレクトリの名前に拡張子をつけるかどうか
    // 拡張子違いのブックがある場合は true にする
    public bool BookDirExt { get; set; } = false;
    public string Encoding { get; set; } = "utf-8";
}
