using System.Text.RegularExpressions;
using OrderHelperWinForms.Models;

namespace OrderHelperWinForms.Services;

public static class ValidationService
{
    static readonly Regex NumberPattern = new(@"^\d+(\.\d+)?$", RegexOptions.Compiled);
    static readonly Regex IntegerPattern = new(@"^\d+$", RegexOptions.Compiled);
    static readonly Regex PhoneFaxPattern = new(@"^[0-9#()\-.\s+]+$", RegexOptions.Compiled);
    static readonly Regex TaxIdPattern = new(@"^\d{8}$", RegexOptions.Compiled);

    static string FieldValue(OrderRow row, string field) => field switch
    {
        "訂購單號" => row.OrderNo,
        "品名規格" => row.Name,
        "廠商名稱" => row.Vendor,
        "訂購量"   => row.Quantity,
        "計價單位" => row.Unit,
        "料號"     => row.Code,
        "料號項次" => row.ItemNo,
        "電話"     => row.Tel,
        "傳真"     => row.Fax,
        _          => "",
    };

    public static List<ValidationError> Validate(List<OrderRow> rows, ValidationConfig config)
    {
        var errors = new List<ValidationError>();
        var activeRules = config.Rules.Where(r => r.Enabled).ToList();

        for (int i = 0; i < rows.Count; i++)
        {
            foreach (var rule in activeRules)
            {
                string val = FieldValue(rows[i], rule.Field);
                bool fail = rule.RuleType switch
                {
                    RuleType.Required  => string.IsNullOrWhiteSpace(val),
                    RuleType.Regex     => !string.IsNullOrEmpty(rule.Parameter)
                                         && !string.IsNullOrWhiteSpace(val)
                                         && !SafeRegexMatch(val, rule.Parameter),
                    RuleType.MaxLength => int.TryParse(rule.Parameter, out int max)
                                         && val.Length > max,
                    RuleType.Number => !IsBlank(val) && !NumberPattern.IsMatch(val),
                    RuleType.PositiveNumber => !IsBlank(val) && (!decimal.TryParse(val, out var pn) || pn <= 0),
                    RuleType.Integer => !IsBlank(val) && !IntegerPattern.IsMatch(val),
                    RuleType.PositiveInteger => !IsBlank(val) && (!int.TryParse(val, out var pi) || pi <= 0),
                    RuleType.NotZero => !IsBlank(val) && (!decimal.TryParse(val, out var nz) || nz == 0),
                    RuleType.PhoneFax => !IsBlank(val) && !PhoneFaxPattern.IsMatch(val),
                    RuleType.TaxId8 => !IsBlank(val) && !TaxIdPattern.IsMatch(val),
                    RuleType.MaxLength20 => val.Length > 20,
                    RuleType.MaxLength50 => val.Length > 50,
                    RuleType.MaxLength100 => val.Length > 100,
                    _                  => false,
                };
                if (fail)
                    errors.Add(new ValidationError(i, rows[i].OrderNo, rule.Field, EffectiveMessage(rule)));
            }
        }
        return errors;
    }

    static bool IsBlank(string value) => string.IsNullOrWhiteSpace(value);

    static bool SafeRegexMatch(string value, string pattern)
    {
        try { return Regex.IsMatch(value, pattern); }
        catch (ArgumentException) { return false; }
    }

    static string EffectiveMessage(ValidationRule rule)
        => string.IsNullOrWhiteSpace(rule.Message)
            ? DefaultMessage(rule.Field, rule.RuleType, rule.Parameter)
            : rule.Message;

    public static string DefaultMessage(string field, RuleType ruleType, string parameter = "")
        => ruleType switch
        {
            RuleType.Required => $"{field}不可空白",
            RuleType.Number => $"{field}必須是數字",
            RuleType.PositiveNumber => $"{field}必須大於 0",
            RuleType.Integer => $"{field}必須是整數",
            RuleType.PositiveInteger => $"{field}必須是大於 0 的整數",
            RuleType.NotZero => $"{field}不可為 0",
            RuleType.PhoneFax => $"{field}只能包含數字、空白與常用電話符號",
            RuleType.TaxId8 => $"{field}必須為 8 碼數字",
            RuleType.MaxLength20 => $"{field}不得超過 20 字",
            RuleType.MaxLength50 => $"{field}不得超過 50 字",
            RuleType.MaxLength100 => $"{field}不得超過 100 字",
            RuleType.MaxLength when int.TryParse(parameter, out var max) => $"{field}不得超過 {max} 字",
            _ => $"{field}格式不正確",
        };

    public static readonly string[] KnownFields =
    {
        "訂購單號", "品名規格", "廠商名稱", "訂購量",
        "計價單位", "料號", "料號項次", "電話", "傳真",
    };
}
