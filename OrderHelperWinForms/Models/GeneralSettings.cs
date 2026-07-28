namespace OrderHelperWinForms.Models;

public enum PdfOutputMode
{
    ExcelDirectory,
    DefaultDirectory,
    AskEveryTime,
}

public class GeneralSettings
{
    public string? LastExcelDirectory  { get; set; }
    public bool    AutoSaveSameDir     { get; set; } = true;
    public PdfOutputMode PdfOutputMode { get; set; } = PdfOutputMode.ExcelDirectory;
    public string? DefaultPdfDirectory { get; set; }
    public int     MaxRowsPerPage      { get; set; } = 12;

    public static GeneralSettings Default() => new();
}
