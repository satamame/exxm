using System.Text;
using System.Diagnostics;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Vbe.Interop;

namespace ExcelMacro;

public class ExcelMacroIO
{
    /// <summary>
    /// 指定したディレクトリから Excel ブックを探してファイル名のリストを返す。
    /// </summary>
    /// <param name="dir">Excel ブックがあるディレクトリ</param>
    /// <param name="exclude">除外するファイル名のリスト</param>
    /// <param name="ext">対象とする拡張子</param>
    /// <returns></returns>
    public static List<string> FindExcelFiles(
        string dir, List<string> exclude, List<string> ext)
    {
        var files = new List<string>();

        foreach (var e in ext)
        {
            if (e.StartsWith('.'))
            {
                files.AddRange(Directory.GetFiles(dir, $"*{e}"));
            }
            else
            {
                files.AddRange(Directory.GetFiles(dir, $"*.{e}"));
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
                + "処理を中止します。";
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
            app = (Excel.Application)Marshal2.Marshal2.GetActiveObject(
                "Excel.Application");
        }
        catch (COMException)
        {
            app = new Excel.Application();
            isRunning = false;
        }
        app.Visible = false;
        return (app, isRunning);
    }

    /// <summary>
    /// Excel ブックがすでに開いていればそれを、さもなくば新たに開いて返す。
    /// </summary>
    /// <param name="app">ブックを開く Excel インスタンス</param>
    /// <param name="path">開きたい Excel ブックのパス</param>
    /// <returns>Excel.Workbook wb, bool isOpen</returns>
    public static (Excel.Workbook, bool) GetOrOpenWorkbook(
        Excel.Application app, string path)
    {
        string bookName = Path.GetFileName(path);
        Excel.Workbook wb;
        bool isOpen = true;
        try
        {
            wb = app.Workbooks[bookName];
        }
        catch (COMException)
        {
            // Open メソッドに絶対パスを渡さないとエラーになる (原因不明)
            wb = app.Workbooks.Open(Path.GetFullPath(path));
            isOpen = false;
        }
        return (wb, isOpen);
    }

    public static void ReleaseBook(Excel.Workbook wb, bool isOpen)
    {
        // シートをすべて解放する
        foreach (Excel.Worksheet ws in wb.Worksheets)
        {
            _ = Marshal.ReleaseComObject(ws);
        }

        // ブックが元々開いていたのでなければ閉じる
        if (!isOpen)
        {
            wb.Close(false);
        }
        // ブックを解放する
        _ = Marshal.ReleaseComObject(wb);
    }

    public static void ReleaseApp(Excel.Application app, bool isRunning)
    {
        // Visible に戻す
        app.Visible = true;

        // インスタンスが元々起動していたのでなければ終了する
        if (!isRunning)
        {
            app.Quit();
        }
        // インスタンスを解放する
        _ = Marshal.ReleaseComObject(app);
    }

    /// <summary>
    /// Excel アプリケーション、ブックを解放する。
    /// </summary>
    /// <param name="app">終了する Excel インスタンス</param>
    /// <param name="wb">閉じる Excel ブック</param>
    /// <param name="isRunning">インスタンスが元々起動していたか</param>
    /// <param name="isOpen">ブックが元々開いていたか</param>
    public static void Release(
        Excel.Application app, Excel.Workbook wb, bool isRunning, bool isOpen)
    {
        ReleaseBook(wb, isOpen);
        ReleaseApp(app, isRunning);
    }

    /// <summary>
    /// Excel ブックから VBA マクロを抽出し .bas ファイルとして保存する。
    /// </summary>
    /// <param name="filePath">Excel ブックのパス</param>
    /// <param name="macrosDir">.bas ファイルの保存先ディレクトリ</param>
    /// <param name="clean">この引数は未使用です</param>
    public static void ExtractMacros(string filePath, string macrosDir, bool clean)
    {
        CheckMultipleInstances();
        (var app, var isRunning) = GetExcelInstance();

        Excel.Workbook wb;
        bool isOpen;
        try
        {
            (wb, isOpen) = GetOrOpenWorkbook(app, filePath);
        }
        catch (Exception)
        {
            ReleaseApp(app, isRunning);
            throw;
        }
        string wbName = Path.GetFileName(filePath);
        Console.WriteLine($"{wbName} の処理を開始しました。");

        // TODO: clean オプションを実装する

        // 保存先のディレクトリがなければ作成する。
        // e.g. "Full/path/to/macros/Book1.xlsm"
        string destDir = Path.Combine(Path.GetFullPath(macrosDir), wbName);
        Directory.CreateDirectory(destDir);

        VBComponents vbaComponents = wb.VBProject.VBComponents;

        try
        {
            foreach (VBComponent component in vbaComponents)
            {
                string componentName = component.Name;
                if (component.Type == vbext_ComponentType.vbext_ct_MSForm)
                {
                    // フォームコンポーネントは無視する。
                    continue;
                }
                else if (component.Type == vbext_ComponentType.vbext_ct_StdModule)
                {
                    // 標準モジュール
                    component.Export(Path.Combine(destDir, $"{componentName}.bas"));
                    Console.WriteLine($"{componentName} を抽出しました。");
                }
                else if (component.Type == vbext_ComponentType.vbext_ct_ClassModule)
                {
                    // クラスモジュールは無視する。
                    continue;
                }
                else if (component.Type == vbext_ComponentType.vbext_ct_Document)
                {
                    // ドキュメント（シートなど）
                    string? sheetName = component.Properties.Item("Name").Value.ToString();
                    if (sheetName != null)
                    {
                        componentName = $"{componentName} ({sheetName})";
                    }
                    component.Export(Path.Combine(destDir, $"{componentName}.bas"));
                    Console.WriteLine($"{componentName} を抽出しました。");
                }
            }
        }
        catch (COMException)
        {
            throw;
        }
        finally
        {
            // 例外が起きても起きなくてもブックとアプリケーションを解放する。
            Release(app, wb, isRunning, isOpen);
        }
    }

    /// <summary>
    /// Excel ブックへ VBA マクロを書き戻す。
    /// </summary>
    /// <param name="filePath">Excel ブックのパス</param>
    /// <param name="macrosDir">.bas ファイルの保存先ディレクトリ</param>
    /// <param name="clean">この引数は未使用です</param>
    /// <exception cref="DirectoryNotFoundException"></exception>
    public static void WriteBackMacros(string filePath, string macrosDir, bool clean)
    {
        CheckMultipleInstances();
        (var app, var isRunning) = GetExcelInstance();

        Excel.Workbook wb;
        bool isOpen;
        try
        {
            (wb, isOpen) = GetOrOpenWorkbook(app, filePath);
        }
        catch (Exception)
        {
            ReleaseApp(app, isRunning);
            throw;
        }
        string wbName = Path.GetFileName(filePath);
        Console.WriteLine($"{wbName} の処理を開始しました。");

        // .bas ファイルが保存されているディレクトリ
        var basDir = Path.Combine(Path.GetFullPath(macrosDir), wbName);

        // ディレクトリが存在しなければ例外をスローする。
        if (!Directory.Exists(basDir))
        {
            Release(app, wb, isRunning, isOpen);
            throw new DirectoryNotFoundException(
                $"ディレクトリが見つかりません: {basDir}");
        }

        // basDir にある .bas ファイルのリストを取得する。
        var basFiles = Directory.GetFiles(basDir, "*.bas");

        foreach (var basFile in basFiles)
        {
            // basFile に書かれた Attribute VB_Name からコンポーネント名を取得する。
            string componentName = "";
            Encoding shiftJisEncoding = Encoding.GetEncoding("shift_jis");
            using (var sr = new StreamReader(basFile, shiftJisEncoding))
            {
                string? line;
                while ((line = sr.ReadLine()) != null)
                {
                    if (line.StartsWith("Attribute VB_Name"))
                    {
                        componentName = line.Split('\"')[1];
                        break;
                    }
                }
            }

            VBComponents vbaComponents = wb.VBProject.VBComponents;

            // 書き戻し先のコンポーネント
            VBComponent? component;
            try
            {
                component = vbaComponents.Item(componentName);
            }
            catch (IndexOutOfRangeException)
            {
                component = null;
            }

            if (component != null && component.Type == vbext_ComponentType.vbext_ct_Document)
            {
                // このコンポーネントはシートまたはブックに関連付けられている。
                // ドキュメントのマクロを上書きする。
                OverwriteDocumentMacro(component, basFile);
            }
            else
            {
                // このコンポーネントはシートまたはブックに関連付けられていない。
                if (component != null)
                {
                    // 既存のモジュールを削除する。
                    vbaComponents.Remove(vbaComponents.Item(componentName));
                }
                // VBA マクロをブックにインポートする。
                wb.VBProject.VBComponents.Import(basFile);
            }
            Console.WriteLine($"{componentName} を書き戻しました。");
        }
        // Excel ブックを保存する。
        wb.Save();

        // ブックとアプリケーションを解放する。
        Release(app, wb, isRunning, isOpen);
    }

    public static void OverwriteDocumentMacro(VBComponent component, string basFile)
    {
        // 既存のコードを削除する。
        CodeModule codeModule = component.CodeModule;
        int lineCount = codeModule.CountOfLines;
        if (lineCount > 0)
        {
            codeModule.DeleteLines(1, lineCount);
        }

        // Shift JIS エンコーディングでファイルの内容を取得する
        Encoding shiftJisEncoding = Encoding.GetEncoding("shift_jis");
        var lines = File.ReadAllLines(basFile, shiftJisEncoding);

        // メタデータの行をスキップする
        bool isMetaData = true;
        var filteredLines = new List<string>();
        foreach (var line in lines)
        {
            if (isMetaData)
            {
                // メタデータのパターンに当てはまるかチェックする
                string trimmedLine = line.TrimStart();
                if (trimmedLine.StartsWith("VERSION")
                    || trimmedLine.StartsWith("BEGIN")
                    || trimmedLine.StartsWith("MultiUse")
                    || trimmedLine.StartsWith("END")
                    || trimmedLine.StartsWith("Attribute"))
                {
                    continue;
                }
                else
                {
                    isMetaData = false;
                }
            }
            filteredLines.Add(line);
        }

        // 残った行を結合する
        string newCode = string.Join(Environment.NewLine, filteredLines);

        // 新しいコードを追加する
        codeModule.AddFromString(newCode);
    }
}
