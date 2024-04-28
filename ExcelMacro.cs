using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Vbe.Interop;
using System.Diagnostics;
using System.Runtime.InteropServices;
using Marshal;

namespace ExcelMacro;

public class ExcelMacroIO
{
    /* 指定したディレクトリから Excel ブックを探してファイル名のリストを返す */
    public static List<string> FindExcelFiles(
        string dir, bool recursive, List<string> exclude, List<string> ext)
    {
        var files = new List<string>();
        var searchOption = recursive
            ? SearchOption.AllDirectories : SearchOption.TopDirectoryOnly;

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

        // TODO: "~$" で始まるファイルを除外する

        return files;
    }

    public static bool ExtractMacros(string path)
    {
        if (ExcelMacroIO.IsMultipleExcelInstancesRunning())
        {
            var msg = "Excel のインスタンスが複数起動しています。\n"
                + "起動するインスタンスは1個までにしてください。\n"
                + "マクロの抽出を中止します。";
            throw new Exception(msg);
        }
        // TODO: エラーハンドリングをして Excel のインスタンスが残らないようにする

        var app = GetExcelInstance();

        // TODO: Workbooks の中に開きたいブックがあれば Open せずにそれを使う

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
        return true;
    }

    /// <summary>
    /// Excel のインスタンスが複数起動しているかどうかを返す
    /// </summary>
    /// <returns>複数起動しているなら true</returns>
    public static bool IsMultipleExcelInstancesRunning()
    {
        Process[] excelProcesses = Process.GetProcessesByName("EXCEL");
        return excelProcesses.Length > 1;
    }

    /// <summary>
    /// 起動中または新規の Excel インスタンスを返す
    /// </summary>
    /// <returns>Excel Application</returns>
    public static Excel.Application GetExcelInstance()
    {
        Excel.Application app;
        try
        {
            app = (Excel.Application)Marshal2.GetActiveObject("Excel.Application");
        }
        catch (COMException)
        {
            app = new Excel.Application();
        }
        return app;
    }
}
