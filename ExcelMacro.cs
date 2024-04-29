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
            }
            else
            {
                files.AddRange(Directory.GetFiles(dir, $"*.{e}", searchOption));
            }
        }

        // exclude に指定されたファイル名と "~$" で始まるファイル名を除外する。
        foreach (var e in exclude)
        {
            files.RemoveAll(f => Path.GetFileName(f) == e);
        }
        files.RemoveAll(f => Path.GetFileName(f).StartsWith("~$"));

        return files;
    }

    public static bool ExtractMacros(string path, bool clean)
    {
        CheckMultipleInstances();
        (var app, var isRunning) = GetExcelInstance();
        (var wb, var isOpen) = GetOrOpenWorkbook(app, path);

        // TODO: エラーハンドリングをして Excel のインスタンスが残らないようにする

        // wb のシート数を取得
        int sheetCount = wb.Sheets.Count;
        Console.WriteLine($"シート数: {sheetCount}");

        var vbaProject = wb.VBProject;
        VBComponents vbaComponents = vbaProject.VBComponents;

        foreach (VBComponent component in vbaComponents)
        {
            Console.WriteLine(component.Name);
        }

        if (!isOpen) wb.Close();
        if (!isRunning) app.Quit();
        return true;
    }

    /// <summary>
    /// Excel のインスタンスが複数起動していれば例外をスローする。
    /// </summary>
    public static void CheckMultipleInstances()
    {
        Process[] excelProcesses = Process.GetProcessesByName("EXCEL");
        if (excelProcesses.Length > 1)
        {
            var msg = "Excel のインスタンスが複数起動しています。\n"
                + "起動するインスタンスは1個までにしてください。\n"
                + "マクロの抽出を中止します。";
            throw new Exception(msg);
        }
    }

    /// <summary>
    /// 起動中または新規の Excel インスタンスを返す
    /// </summary>
    /// <returns>Excel Application app, bool isRunning</returns>
    public static (Excel.Application, bool) GetExcelInstance()
    {
        Excel.Application app;
        bool isRunning = true;
        try
        {
            app = (Excel.Application)Marshal2.GetActiveObject("Excel.Application");
        }
        catch (COMException)
        {
            app = new Excel.Application();
            isRunning = false;
        }
        return (app, isRunning);
    }

    /// <summary>
    /// ブック名を指定してすでに開いていればそれを、さもなくば開いて返す。
    /// </summary>
    /// <returns>Excel.Workbook wb, bool isOpen</returns>
    public static (Excel.Workbook, bool) GetOrOpenWorkbook(Excel.Application app, string path)
    {
        string name = Path.GetFileName(path);
        Excel.Workbook wb;
        bool isOpen = true;
        try
        {
            wb = app.Workbooks[name];
        }
        catch (COMException)
        {
            // Open メソッドに絶対パスを渡さないとエラーになる (原因不明)
            wb = app.Workbooks.Open(Path.GetFullPath(path));
            isOpen = false;
        }
        return (wb, isOpen);
    }
}
