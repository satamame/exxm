using System.Text;
using System.Diagnostics;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Vbe.Interop;
using Settings;

namespace ExcelMacro;

public class MacroIO
{
    public Excel.Application? App { get; set; } = null;
    public Excel.Workbook? Wb { get; set; } = null;
    public List<string> Files { get; set; } = [];
    public AppSettings Settings { get; set; }
    private bool AppRunning = false;
    private bool WbOpen = false;

    /// <summary>
    /// コンストラクタ
    /// </summary>
    /// <param name="settings"></param>
    public MacroIO(AppSettings settings)
    {
        this.Settings = settings;

        // 拡張子の先頭に "." がなければつける。
        var ext = this.Settings.Excel.Ext;
        for (int i = 0; i < ext.Count; i++)
        {
            if (!ext[i].StartsWith('.'))
            {
                ext[i] = "." + ext[i];
            }
        }

        // 対象となる Excel ブックのファイル名を取得する。
        this.Files = FindExcelFiles();
    }

    /// <summary>
    /// 対象となる Excel ブックのファイル名のリストを返す関数
    /// </summary>
    /// <returns>ファイル名のリスト</returns>
    protected List<string> FindExcelFiles()
    {
        var files = new List<string>();
        var dir = this.Settings.Excel.Dir;
        var ext = this.Settings.Excel.Ext;
        var exclude = this.Settings.Excel.Exclude;

        foreach (var e in ext)
        {
            files.AddRange(Directory.GetFiles(dir, $"*{e}"));
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
    /// Excel のインスタンスが複数起動していれば例外をスローする関数
    /// </summary>
    protected static void CheckMultipleInstances()
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
    /// 起動中または新規の Excel インスタンスを返す関数
    /// </summary>
    /// <returns>Excel Application app, bool appRunning</returns>
    protected static (Excel.Application, bool) GetExcelInstance()
    {
        Excel.Application app;
        bool appRunning = true;
        try
        {
            app = (Excel.Application)Marshal2.Marshal2.GetActiveObject(
                "Excel.Application");
        }
        catch (COMException)
        {
            app = new Excel.Application();
            appRunning = false;
        }
        app.Visible = false;
        return (app, appRunning);
    }

    /// <summary>
    /// Excel ブックがすでに開いていればそれを、さもなくば新たに開いて返す関数
    /// </summary>
    /// <param name="app">ブックを開く Excel インスタンス</param>
    /// <param name="path">開きたい Excel ブックのパス</param>
    /// <returns>Excel.Workbook wb, bool wbOpen</returns>
    protected static (Excel.Workbook, bool) GetOrOpenWorkbook(
        Excel.Application app, string path)
    {
        string bookName = Path.GetFileName(path);
        Excel.Workbook wb;
        bool wbOpen = true;
        try
        {
            wb = app.Workbooks[bookName];
        }
        catch (COMException)
        {
            // Open メソッドに絶対パスを渡さないとエラーになる (原因不明)
            wb = app.Workbooks.Open(Path.GetFullPath(path));
            wbOpen = false;
        }
        return (wb, wbOpen);
    }

    protected void ReleaseBook()
    {
        if (this.Wb == null) return;

        // シートをすべて解放する
        foreach (Excel.Worksheet ws in this.Wb.Worksheets)
        {
            _ = Marshal.ReleaseComObject(ws);
        }

        // ブックが元々開いていたのでなければ閉じる。
        if (!this.WbOpen)
        {
            this.Wb.Close(false);
        }
        // ブックを解放する。
        _ = Marshal.ReleaseComObject(this.Wb);
        this.Wb = null;
    }

    protected void ReleaseApp()
    {
        if (this.App == null) return;

        // インスタンスが元々起動していたのでなければ終了する。
        if (!this.AppRunning)
        {
            this.App.Quit();
        }
        else
        {
            this.App.Visible = true;
        }

        // インスタンスを解放する
        _ = Marshal.ReleaseComObject(this.App);
        this.App = null;
    }

    /// <summary>
    /// Excel アプリケーション、ブックを解放する関数
    /// </summary>
    protected void Release()
    {
        this.ReleaseBook();
        this.ReleaseApp();
    }

    /// <summary>
    /// Excel ブックから VBA マクロを抽出し .bas ファイルとして保存する関数
    /// </summary>
    /// <param name="filePath">Excel ブックのパス</param>
    /// <param name="clean">この引数は未使用です</param>
    public void ExtractMacros(string filePath, bool clean)
    {
        MacroIO.CheckMultipleInstances();
        (this.App, this.AppRunning) = MacroIO.GetExcelInstance();

        try
        {
            (this.Wb, this.WbOpen) = MacroIO.GetOrOpenWorkbook(this.App, filePath);
        }
        catch (Exception)
        {
            this.ReleaseApp();
            throw;
        }
        string wbName = Path.GetFileName(filePath);
        Console.WriteLine($"{wbName} の処理を開始しました。");

        // TODO: clean オプションを実装する

        // 保存先のディレクトリがなければ作成する。
        // e.g. "Full/path/to/macros/Book1.xlsm"
        string bookDir = wbName;
        if (!this.Settings.Macros.BookDirExt)
        {
            bookDir = Path.GetFileNameWithoutExtension(wbName);
        }
        string destDir = Path.Combine(
            Path.GetFullPath(this.Settings.Macros.Dir), bookDir);
        Directory.CreateDirectory(destDir);

        VBComponents vbaComponents = this.Wb.VBProject.VBComponents;

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
            Release();
        }
    }

    /// <summary>
    /// Excel ブックへ VBA マクロを書き戻す。
    /// </summary>
    /// <param name="filePath">Excel ブックのパス</param>
    /// <param name="macrosDir">.bas ファイルの保存先ディレクトリ</param>
    /// <param name="clean">この引数は未使用です</param>
    /// <exception cref="DirectoryNotFoundException"></exception>
    public void WriteBackMacros(string filePath, bool clean)
    {
        MacroIO.CheckMultipleInstances();
        (this.App, this.AppRunning) = MacroIO.GetExcelInstance();

        try
        {
            (this.Wb, this.WbOpen) = MacroIO.GetOrOpenWorkbook(this.App, filePath);
        }
        catch (Exception)
        {
            this.ReleaseApp();
            throw;
        }
        string wbName = Path.GetFileName(filePath);
        Console.WriteLine($"{wbName} の処理を開始しました。");

        // .bas ファイルが保存されているディレクトリ
        string bookDir = wbName;
        if (!this.Settings.Macros.BookDirExt)
        {
            bookDir = Path.GetFileNameWithoutExtension(wbName);
        }
        var basDir = Path.Combine(
            Path.GetFullPath(this.Settings.Macros.Dir), bookDir);

        // ディレクトリが存在しなければ例外をスローする。
        if (!Directory.Exists(basDir))
        {
            Release();
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

            VBComponents vbaComponents = this.Wb.VBProject.VBComponents;

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
                this.Wb.VBProject.VBComponents.Import(basFile);
            }
            Console.WriteLine($"{componentName} を書き戻しました。");
        }
        // Excel ブックを保存する。
        this.Wb.Save();

        // ブックとアプリケーションを解放する。
        Release();
    }

    protected static void OverwriteDocumentMacro(VBComponent component, string basFile)
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
