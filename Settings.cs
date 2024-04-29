namespace Settings;

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
}
