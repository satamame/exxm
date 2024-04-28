using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Vbe.Interop;

namespace ExcelMacro;

public class ExcelMacroIO
{
    /* 指定したディレクトリから Excel ブックを探してファイル名のリストを返す */
    public List<string> FindExcelFiles(string dir, bool recursive, List<string> exclude, List<string> ext)
    {
        var files = new List<string>();
        var searchOption = recursive ? SearchOption.AllDirectories : SearchOption.TopDirectoryOnly;

        foreach (var e in ext)
        {
            if (e.StartsWith('.'))
            {
                files.AddRange(Directory.GetFiles(dir, $"*{e}", searchOption));
            } else {
                files.AddRange(Directory.GetFiles(dir, $"*.{e}", searchOption));
            }
        }

        foreach (var e in exclude)
        {
            files.RemoveAll(f => f.Contains(e));
        }

        return files;
    }

    public void ReadMacro(string path)
    {
        // TODO: エラーハンドリングをして Excel のインスタンスが残らないようにする

        var app = new Excel.Application();
        // Open メソッドに絶対パスを渡さないとエラーになる (原因不明)
        var wb = app.Workbooks.Open(Path.GetFullPath(path));

        // wb のシート数を取得
        int sheetCount = wb.Sheets.Count;
        Console.WriteLine($"シート数: {sheetCount}");


        var vbaProject = wb.VBProject;
        VBComponents vbaComponents = vbaProject.VBComponents;

        foreach (VBComponent component in vbaComponents)
        {
            Console.WriteLine(component.Name);
        }

        wb.Close();
        app.Quit();
    }
}
