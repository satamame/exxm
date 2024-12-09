using System.Text;
using System.Diagnostics;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Vbe.Interop;
using Settings;

namespace ExcelMacro;

public class MacroIO
{
    private AppSettings Settings { get; set; }
    private Excel.Application? App { get; set; } = null;
    private Excel.Workbook? Wb { get; set; } = null;
    private bool AppWasRunning { get; set; } = false;
    private bool WbWasOpen { get; set; } = false;
    public List<string> WbFiles { get; set; } = new List<string>();

    /// <summary>
    /// コンストラクタ
    /// </summary>
    /// <param name="settings">アプリの設定</param>
    /// <param name="target">1個のファイルを指定する場合のファイル名</param>
    public MacroIO(AppSettings settings, string target)
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

        if (target != "")
        {
            this.WbFiles = [Path.Combine(settings.Excel.Dir, target)];
        }
        else
        {
            // 対象となる Excel ブックのファイル名を取得する。
            this.WbFiles = this.FindWbFiles();
        }
    }

    /// <summary>
    /// 対象となる Excel ブックのファイル名のリストを返す関数
    /// </summary>
    /// <returns>ファイル名のリスト</returns>
    protected List<string> FindWbFiles()
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

        // DEBUG: 起動中の Excel インスタンスの数を表示する
        Console.WriteLine($"起動中の Excel インスタンス数: {excelProcesses.Length}");

        if (excelProcesses.Length > 1)
        {
            var msg = "Excel のインスタンスが複数起動しています。\n"
                + "起動するインスタンスは1個までにしてください。\n"
                + "処理を中止します。";
            throw new Exception(msg);
        }
    }

    /// <summary>
    /// 起動した (またはすでに起動している) Excel インスタンスを保持する関数
    /// </summary>
    protected void SetupApp()
    {
        try
        {
            // すでに起動している Excel インスタンスを取得する。
            // 複数のインスタンスが起動していないことを確認済みであること。
            this.App = (Excel.Application)Marshal2.Marshal2.GetActiveObject(
                "Excel.Application");
            this.AppWasRunning = true;
        }
        catch (COMException)
        {
            // Excel インスタンスが起動していない場合は新たに起動する。
            this.App = new Excel.Application();
            this.AppWasRunning = false;
        }
    }

    /// <summary>
    /// 開いた (またはすでに開いている) Excel ブックを保持する関数
    /// </summary>
    /// <param name="path">Excel ブックのパス</param>
    protected void SetupWb(string path)
    {
        if (this.App == null)
        {
            throw new Exception("Excel アプリケーションを起動できません。");
        }

        string fullPath = Path.GetFullPath(path);

        // パスが fullPath であるブックが開いていれば保持する。
        foreach (Excel.Workbook wb in this.App.Workbooks)
        {
            if (wb.FullName.Equals(fullPath, StringComparison.OrdinalIgnoreCase))
            {
                this.Wb = wb;
                this.WbWasOpen = true;
                return;
            }
        }

        // 同じ名前のブックが開いているか確認する。
        string wbName = Path.GetFileName(path);
        Excel.Workbook? tmpWb;
        bool sameNameExists;
        try
        {
            tmpWb = this.App.Workbooks[wbName];
            sameNameExists = true;
            Marshal.ReleaseComObject(tmpWb);
        }
        catch (COMException)
        {
            // 例外をキャッチしたので、同じ名前のブックは開いていない。
            sameNameExists = false;
        }

        if (sameNameExists)
        {
            // 同じ名前のブックが開いていた場合はエラーにする。
            throw new Exception($"{wbName} と同じ名前のブックが開いています。");
        }

        // パスが fullPath であるブックを開いて保持する。
        this.Wb = this.App.Workbooks.Open(fullPath);
        this.WbWasOpen = false;
    }

    /// <summary>
    /// Excel ブックを解放する関数
    /// </summary>
    protected void ReleaseWb()
    {
        if (this.Wb == null) return;

        // シートをすべて解放する。
        foreach (Excel.Worksheet ws in this.Wb.Worksheets)
        {
            Marshal.FinalReleaseComObject(ws);
        }
        Marshal.FinalReleaseComObject(this.Wb.Worksheets);

        // VBProject が利用している外部オブジェクトへの参照を解放する。
        References refs = this.Wb.VBProject.References;
        for (int i = refs.Count; i > 0; i--)
        {
            if (!refs.Item(i).BuiltIn)
            {
                refs.Remove(refs.Item(i));
            }
        }

        // VBProject を解放する
        foreach (VBComponent component in this.Wb.VBProject.VBComponents)
        {
            Marshal.FinalReleaseComObject(component);
        }
        Marshal.FinalReleaseComObject(this.Wb.VBProject.VBComponents);
        Marshal.FinalReleaseComObject(this.Wb.VBProject);

        // ブックが元々開いていたのでなければ閉じる。
        if (!this.WbWasOpen)
        {
            this.Wb.Close(false);
        }

        // ブックを解放する。
        Marshal.FinalReleaseComObject(this.Wb);
        this.Wb = null;

        // ガベージコレクションを強制実行する。
        GC.Collect();
        GC.WaitForPendingFinalizers();
    }

    /// <summary>
    /// Excel インスタンスを解放する関数
    /// </summary>
    protected void ReleaseApp()
    {
        if (this.App == null) return;

        // ブックコレクションを解放する。
        Marshal.FinalReleaseComObject(this.App.Workbooks);

        // インスタンスが元々起動していたのでなければ終了する。
        if (!this.AppWasRunning)
        {
            IntPtr hWnd = this.App.Hwnd;
            this.App.Quit();

            // アプリを解放する。
            Marshal.FinalReleaseComObject(this.App);

            // バックグラウンドで起動したままになっていたら強制終了する。
            Process[] excelProcesses = Process.GetProcessesByName("EXCEL");
            foreach (var process in excelProcesses)
            {
                if (process.MainWindowHandle == hWnd)
                {
                    try
                    {
                        process.Kill();
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"エラー: {ex.Message}");
                    }
                }
            }
        }
        else
        {
            // アプリを解放する。
            Marshal.FinalReleaseComObject(this.App);
        }

        this.App = null;

        // ガベージコレクションを強制実行する。
        GC.Collect();
        GC.WaitForPendingFinalizers();
    }

    /// <summary>
    /// エンコーディングを指定して VB コンポーネントをエクスポートする関数
    /// </summary>
    /// <param name="component">VB コンポーネント</param>
    /// <param name="path">エクスポート先のパス</param>
    /// <param name="encoding">保存時のエンコーディング</param>
    protected static void ExportComponent(
        VBComponent component, string path, Encoding encoding)
    {
        // VBAコードをエクスポート
        component.Export(path);

        // エンコーディングの設定が shift_jis でなければ
        // エンコーディングを変えて保存しなおす。
        var sjisEncoding = Encoding.GetEncoding("shift_jis");
        if (encoding != sjisEncoding)
        {
            // ファイルを開き、内容を読み取る。
            string code = File.ReadAllText(path, sjisEncoding);
            // 内容を指定したエンコーディングで再保存する。
            File.WriteAllText(path, code, encoding);
        }
    }

    /// <summary>
    /// Excel ブックから VBA マクロを抽出し .bas ファイルとして保存する関数
    /// </summary>
    /// <param name="filePath">Excel ブックのパス</param>
    /// <param name="clean">最初に保存先をクリアするか (未実装)</param>
    protected void ExtractMacrosFromWb(string filePath, bool clean)
    {
        if (this.App == null)
        {
            throw new Exception("Excel アプリケーションを起動できません。");
        }
        this.SetupWb(filePath);
        VBComponents vbaComponents = this.Wb!.VBProject.VBComponents;

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

        // 保存時に使うエンコーディング
        var encoding = this.Settings.Macros.GetEncodingObj();

        try
        {
            foreach (VBComponent component in vbaComponents)
            {
                string componentName = component.Name;

                // コードがないコンポーネントは無視する。
                CodeModule codeModule = component.CodeModule;
                int modLineCount = codeModule.CountOfLines;
                if (modLineCount < 1)
                {
                    continue;
                }
                string code = codeModule.Lines[1, modLineCount].Trim();
                if (code == "" || code == "Option Explicit")
                {
                    continue;
                }

                if (component.Type == vbext_ComponentType.vbext_ct_MSForm)
                {
                    // フォームコンポーネントは無視する。
                    continue;
                }
                else if (component.Type == vbext_ComponentType.vbext_ct_StdModule)
                {
                    // 標準モジュール
                    var destPath = Path.Combine(destDir, $"{componentName}.bas");
                    MacroIO.ExportComponent(component, destPath, encoding);
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
                    var destPath = Path.Combine(destDir, $"{componentName}.bas");
                    MacroIO.ExportComponent(component, destPath, encoding);
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
            // 例外が起きても起きなくてもブックを解放する。
            this.ReleaseWb();
        }
    }

    /// <summary>
    /// アプリの設定に従って VBA マクロを抽出する関数
    /// </summary>
    /// <param name="clean">最初に保存先をクリアするか (未実装)</param>
    public void ExtractMacros(bool clean)
    {
        MacroIO.CheckMultipleInstances();
        this.SetupApp();

        foreach (var f in this.WbFiles)
        {
            try
            {
                this.ExtractMacrosFromWb(f, clean);
            }
            catch (Exception e)
            {
                Console.WriteLine($"エラー: {e.Message}");
                this.ReleaseWb();
                break;
            }
        }

        this.ReleaseApp();
    }

    /// <summary>
    /// ドキュメントのマクロを上書きする関数
    /// </summary>
    /// <param name="component">VB コンポーネント</param>
    /// <param name="basFile">.bas ファイルのパス</param>
    /// <param name="encoding">.bas ファイルのエンコーディング</param>
    protected static void OverwriteDocumentMacro(
        VBComponent component, string basFile, Encoding encoding)
    {
        // 既存のコードを削除する。
        CodeModule codeModule = component.CodeModule;
        int lineCount = codeModule.CountOfLines;
        if (lineCount > 0)
        {
            codeModule.DeleteLines(1, lineCount);
        }

        var lines = File.ReadAllLines(basFile, encoding);

        // メタデータの行をスキップする。
        bool isMetaData = true;
        var filteredLines = new List<string>();
        foreach (var line in lines)
        {
            if (isMetaData)
            {
                // メタデータのパターンに当てはまるかチェックする。
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

        // 残った行を結合する。
        string newCode = string.Join(Environment.NewLine, filteredLines);

        // 新しいコードを追加する。
        codeModule.AddFromString(newCode);
    }

    /// <summary>
    /// VBA マクロをブックにインポートする関数
    /// </summary>
    /// <param name="vbaComponents">ブックのコンポーネントコレクション</param>
    /// <param name="path">.bas ファイルのパス</param>
    /// <param name="encoding">.bas ファイルのエンコーディング</param>
    protected static void ImportComponent(
               VBComponents vbaComponents, string path, Encoding encoding)
    {
        var sjisEncoding = Encoding.GetEncoding("shift_jis");
        if (encoding != sjisEncoding)
        {
            // エンコーディングの設定が shift_jis でなければ
            // shift_jis で一時ファイルを作ってそれをインポートする。
            string code = File.ReadAllText(path, encoding);
            string tempFilePath = Path.GetTempFileName();
            File.WriteAllText(tempFilePath, code, sjisEncoding);
            vbaComponents.Import(tempFilePath);
            File.Delete(tempFilePath);
        }
        else
        {
            // エンコーディングの設定が shift_jis ならそのままインポートする。
            vbaComponents.Import(path);
        }
    }

    /// <summary>
    /// Excel ブックへ VBA マクロを書き戻す関数
    /// </summary>
    /// <param name="filePath">Excel ブックのパス</param>
    /// <param name="clean">最初にマクロをクリアするか (未実装)</param>
    /// <exception cref="DirectoryNotFoundException"></exception>
    protected void WriteMacrosToWb(string filePath, bool clean)
    {
        if (this.App == null)
        {
            throw new Exception("Excel アプリケーションを起動できません。");
        }
        this.SetupWb(filePath);
        VBComponents vbaComponents = this.Wb!.VBProject.VBComponents;

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
            ReleaseWb();
            throw new DirectoryNotFoundException(
                $"ディレクトリが見つかりません: {basDir}");
        }

        // TODO: clean オプションを実装する

        // basDir にある .bas ファイルのリストを取得する。
        var basFiles = Directory.GetFiles(basDir, "*.bas");

        // ファイルの読み込みに使うエンコーディング
        Encoding encoding = this.Settings.Macros.GetEncodingObj();

        foreach (var basFile in basFiles)
        {
            // basFile に書かれた Attribute VB_Name からコンポーネント名を取得する。
            string componentName = "";
            using (var sr = new StreamReader(basFile, encoding))
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

            if (component != null
                && component.Type == vbext_ComponentType.vbext_ct_Document)
            {
                // このコンポーネントはシートまたはブックに関連付けられている。
                // ドキュメントのマクロを上書きする。
                MacroIO.OverwriteDocumentMacro(component, basFile, encoding);
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
                MacroIO.ImportComponent(vbaComponents, basFile, encoding);
            }
            Console.WriteLine($"{componentName} を書き戻しました。");
        }
        // Excel ブックを保存する。
        this.Wb.Save();

        // ブックを解放する。
        this.ReleaseWb();
    }

    /// <summary>
    /// アプリの設定に従って VBA マクロを書き戻す関数
    /// </summary>
    /// <param name="clean">最初にマクロをクリアするか (未実装)</param>
    public void WriteMacros(bool clean)
    {
        MacroIO.CheckMultipleInstances();
        this.SetupApp();

        foreach (var f in this.WbFiles)
        {
            try
            {
                this.WriteMacrosToWb(f, clean);
            }
            catch (Exception e)
            {
                Console.WriteLine($"エラー: {e.Message}");
                this.ReleaseWb();
                break;
            }
        }

        this.ReleaseApp();
    }
}
