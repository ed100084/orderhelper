using System.Text.RegularExpressions;
using OrderHelperWinForms.Models;

namespace OrderHelperWinForms.Services;

public static class ValidationService
{
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
                                         && !Regex.IsMatch(val, rule.Parameter),
                    RuleType.MaxLength => int.TryParse(rule.Parameter, out int max)
                                         && val.Length > max,
                    _                  => false,
                };
                if (fail)
                    errors.Add(new ValidationError(i, rows[i].OrderNo, rule.Field, rule.Message));
            }
        }
        return errors;
    }

    public static readonly string[] KnownFields =
    {
        "訂購單號", "品名規格", "廠商名稱", "訂購量",
        "計價單位", "料號", "料號項次", "電話", "傳真",
    };
}
