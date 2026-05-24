using System.Text.RegularExpressions;
using OrderHelperWinForms.Models;

namespace OrderHelperWinForms.Services;

public static partial class TextHelper
{
    static readonly Dictionary<string, string> Replacements = new()
    {
        // Windows PMingLiU contains U+9039 but bundled Noto Sans TC does not.
        ["逹"] = "達",
    };

    public static string NormalizeText(object? value)
    {
        if (value is null) return "";
        // Float integers become ints (e.g. 3.0 → "3")
        if (value is double d && d == Math.Floor(d))
            return ((long)d).ToString();
        if (value is float f && f == MathF.Floor(f))
            return ((long)f).ToString();
        string text = value.ToString()?.Trim() ?? "";
        foreach (var (src, dst) in Replacements)
            text = text.Replace(src, dst);
        return text;
    }

    [GeneratedRegex(@"(\d{3})(\d{2})(\d{2})")]
    private static partial Regex RocDatePattern();

    /// <summary>
    /// Infer order date from filename or first order numbers (ROC calendar pattern YYYMMDD).
    /// Falls back to today's date.
    /// </summary>
    public static string InferOrderDate(string filename, IEnumerable<OrderRow> orders)
    {
        var candidates = new[] { filename }.Concat(orders.Take(5).Select(o => o.OrderNo));
        foreach (var text in candidates)
        {
            var m = RocDatePattern().Match(text);
            if (!m.Success) continue;
            int rocYear = int.Parse(m.Groups[1].Value);
            int month   = int.Parse(m.Groups[2].Value);
            int day     = int.Parse(m.Groups[3].Value);
            return $"{rocYear + 1911:D4}-{month:D2}-{day:D2}";
        }
        return DateTime.Today.ToString("yyyy-MM-dd");
    }
}
