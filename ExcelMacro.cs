using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Vbe.Interop;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text;

namespace ExcelMacro;

public class ExcelMacroIO
{
    /// <summary>
    /// 指定したディレクトリから Excel ブックを探してファイル名のリストを返す。
    /// </summary>
    /// <param name="dir"></param>
    /// <param name="exclude"></param>
    /// <param name="ext"></param>
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

        //Console.WriteLine($"Excel のインスタンス数: {excelProcesses.Length}");

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
            app = (Excel.Application)Marshal2.Marshal2.GetActiveObject(
                "Excel.Application");
        }
        catch (COMException)
        {
            app = new Excel.Application();
            app.Visible = false;
            isRunning = false;
        }
        return (app, isRunning);
    }

    /// <summary>
    /// ブックがすでに開いていればそれを、さもなくば新たに開いて返す。
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

    public static void Close(
        Excel.Application app, Excel.Workbook wb, bool isRunning, bool isOpen)
    {
#pragma warning disable CA1416
        foreach (Excel.Worksheet ws in wb.Worksheets)
        {
            _ = Marshal.ReleaseComObject(ws);
        }
        if (!isOpen)
        {
            wb.Close(false);
        }
        _ = Marshal.ReleaseComObject(wb);
        if (!isRunning)
        {
            app.Quit();
        }
        _ = Marshal.ReleaseComObject(app);
#pragma warning restore CA1416
    }

    public static void ExtractMacros(string filePath, string macrosDir, bool clean)
    {
        CheckMultipleInstances();
        (var app, var isRunning) = GetExcelInstance();
        (var wb, var isOpen) = GetOrOpenWorkbook(app, filePath);

        // TODO: エラーハンドリングをして Excel のインスタンスが残らないようにする
        // TODO: clean オプションを実装する

        // wb のシート数を取得
        //int sheetCount = wb.Sheets.Count;
        //Console.WriteLine($"シート数: {sheetCount}");

        // 保存先のディレクトリがなければ作成する。
        // e.g. "Full/path/to/macros/Book1.xlsm"
        string destDir = Path.Combine(
            Path.GetFullPath(macrosDir), Path.GetFileName(filePath));
        Directory.CreateDirectory(destDir);

        VBComponents vbaComponents = wb.VBProject.VBComponents;

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
                //Console.WriteLine(componentName);
                component.Export(Path.Combine(destDir, $"{componentName}.bas"));
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
            }
        }

        // 後処理
        Close(app, wb, isRunning, isOpen);
    }

    public static void WriteBackMacros(string filePath, string macrosDir, bool clean)
    {
        CheckMultipleInstances();
        (var app, var isRunning) = GetExcelInstance();
        (var wb, var isOpen) = GetOrOpenWorkbook(app, filePath);

        // .bas ファイルが保存されているディレクトリ
        var basDir = Path.Combine(
            Path.GetFullPath(macrosDir), Path.GetFileName(filePath));

        // ディレクトリが存在しなければ例外をスローする。
        if (!Directory.Exists(basDir))
        {
            Close(app, wb, isRunning, isOpen);
            throw new DirectoryNotFoundException(
                $"ディレクトリが見つかりません: {basDir}");
        }

        // basDir にある .bas ファイルのリストを取得する。
        var basFiles = Directory.GetFiles(basDir, "*.bas");

        foreach (var basFile in basFiles)
        {
            // basFile に書かれた Attribute VB_Name からコンポーネント名を取得する。
            string componentName = "";
            using (var sr = new StreamReader(basFile))
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
            VBComponent? component = vbaComponents.Item(componentName);

            if (component.Type == vbext_ComponentType.vbext_ct_Document)
            {
                // このコンポーネントはシートまたはブックに関連付けられています。
                OverwriteDocumentMacro(component, basFile);
            }
            else
            {
                // このコンポーネントはシートまたはブックに関連付けられていません。
                // 既存のモジュールを削除する。
                try
                {
                    vbaComponents.Remove(vbaComponents.Item(componentName));
                }
                catch (COMException)
                {
                    // モジュールが存在しない場合は何もしない。
                }
                // モジュールをインポートする。
                wb.VBProject.VBComponents.Import(basFile);
            }
        }

        // 後処理
        Close(app, wb, isRunning, isOpen);
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

        // エンコーディング プロバイダーを登録する
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

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
                    || trimmedLine.StartsWith("Attribute")
                    || string.IsNullOrWhiteSpace(trimmedLine))
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
