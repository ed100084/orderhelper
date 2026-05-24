using System.Text.Json;
using System.Text.Json.Serialization;

namespace OrderHelperWinForms.Services;

public class LogEntry
{
    public string  Timestamp { get; set; } = "";
    public string  User      { get; set; } = "";
    public string  Action    { get; set; } = "";
    public string? Source    { get; set; }
    public string? Output    { get; set; }
    public bool    Success   { get; set; }
    public string? Error     { get; set; }
}

public static class ActivityLogger
{
    static readonly string LogDir = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
        "OrderHelper", "logs");

    static readonly JsonSerializerOptions Opts = new()
    {
        DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull,
    };

    static string TodayLog => Path.Combine(LogDir, $"{DateTime.Today:yyyy-MM-dd}.jsonl");

    public static void Log(string action, string? source = null, string? output = null,
                           bool success = true, string? error = null)
    {
        var entry = new LogEntry
        {
            Timestamp = DateTime.Now.ToString("yyyy-MM-ddTHH:mm:ss"),
            User      = Environment.UserName,
            Action    = action,
            Source    = source,
            Output    = output,
            Success   = success,
            Error     = error,
        };
        try
        {
            Directory.CreateDirectory(LogDir);
            File.AppendAllText(TodayLog, JsonSerializer.Serialize(entry, Opts) + "\n");
        }
        catch { }
    }

    public static List<LogEntry> LoadAll(DateTime? since = null)
    {
        var entries = new List<LogEntry>();
        if (!Directory.Exists(LogDir)) return entries;

        foreach (var file in Directory.GetFiles(LogDir, "*.jsonl")
                                      .OrderByDescending(f => f))
        {
            // Skip files older than the requested date (filename is yyyy-MM-dd.jsonl)
            if (since.HasValue)
            {
                var name = Path.GetFileNameWithoutExtension(file);
                if (DateTime.TryParse(name, out var fileDate) && fileDate.Date < since.Value.Date)
                    continue;
            }

            try
            {
                foreach (var line in File.ReadAllLines(file))
                {
                    if (string.IsNullOrWhiteSpace(line)) continue;
                    var e = JsonSerializer.Deserialize<LogEntry>(line, Opts);
                    if (e != null) entries.Add(e);
                }
            }
            catch { }
        }
        return entries;
    }
}
