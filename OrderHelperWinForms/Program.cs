using OrderHelperWinForms.Forms;
using OrderHelperWinForms.Services;

internal static class Program
{
    [STAThread]
    static void Main(string[] args)
    {
        // CLI mode: any argument triggers headless execution (no GUI window)
        if (args.Length > 0)
        {
            Environment.Exit(CliRunner.Run(args));
            return;
        }

        // GUI mode — explicit [STAThread] required for OLE drag-and-drop
        Application.EnableVisualStyles();
        Application.SetCompatibleTextRenderingDefault(false);
        Application.SetHighDpiMode(HighDpiMode.SystemAware);
        Application.Run(new MainForm());
    }
}
