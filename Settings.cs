using System.Text;

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
    // マクロの保存時のエンコーディング (utf-8, shift_jis)
    public string Encoding { get; set; } = "utf-8";

    public Encoding GetEncodingObj()
    {
        // this.Encoding が "utf-8" または "shift_jis" であること。
        if (this.Encoding == "utf-8")
        {
            return new UTF8Encoding(false);
        }
        else if (this.Encoding == "shift_jis")
        {
            return System.Text.Encoding.GetEncoding("shift_jis");
        }
        else
        {
            throw new Exception("Invalid encoding");
        }
    }
}
