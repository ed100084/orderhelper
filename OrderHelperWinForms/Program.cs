using OrderHelperWinForms.Forms;
using OrderHelperWinForms.Services;

// CLI mode: any argument triggers headless execution (no GUI window)
if (args.Length > 0)
{
    Environment.Exit(CliRunner.Run(args));
    return;
}

// GUI mode
Application.EnableVisualStyles();
Application.SetCompatibleTextRenderingDefault(false);
Application.SetHighDpiMode(HighDpiMode.SystemAware);

Application.Run(new MainForm());
