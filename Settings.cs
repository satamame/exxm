namespace Settings;

public class AppSettings
{
    public ExcelSettings Excel { get; set; }
    public MacrosSettings Macros { get; set; }
}

public class ExcelSettings
{
    public string Dir { get; set; }
    public bool Recursive { get; set; }
    public List<string> Exclude { get; set; }
    public List<string> Ext { get; set; }
}

public class MacrosSettings
{
    public string Dir { get; set; }
}
