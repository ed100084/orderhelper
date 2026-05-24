namespace OrderHelperWinForms.Models;

public class GeneralSettings
{
    public string? LastExcelDirectory  { get; set; }
    public bool    AutoSaveSameDir     { get; set; } = true;
    public string? DefaultPdfDirectory { get; set; }

    public static GeneralSettings Default() => new();
}
