using System.Text.Json.Serialization;

namespace OrderHelperWinForms.Models;

public enum RuleType
{
    Required,
    Regex,       // Legacy only: existing config files may still contain Regex.
    MaxLength,
    Number,
    PositiveNumber,
    Integer,
    PositiveInteger,
    NotZero,
    PhoneFax,
    TaxId8,
    MaxLength20,
    MaxLength50,
    MaxLength100,
}

public class ValidationRule
{
    public bool     Enabled   { get; set; } = true;
    public string   Field     { get; set; } = "";
    public RuleType RuleType  { get; set; } = RuleType.Required;
    public string   Parameter { get; set; } = "";
    public string   Message   { get; set; } = "";
}

public class ValidationConfig
{
    public List<ValidationRule> Rules { get; set; } = new();

    public static ValidationConfig Default() => new()
    {
        Rules = new List<ValidationRule>
        {
            new() { Enabled = true,  Field = "訂購單號", RuleType = RuleType.Required,  Parameter = "",              Message = "訂購單號不可空白" },
            new() { Enabled = true,  Field = "品名規格", RuleType = RuleType.Required,  Parameter = "",              Message = "品名規格不可空白" },
            new() { Enabled = true,  Field = "廠商名稱", RuleType = RuleType.Required,  Parameter = "",              Message = "廠商名稱不可空白" },
            new() { Enabled = true,  Field = "訂購量",   RuleType = RuleType.PositiveNumber, Parameter = "", Message = "訂購量必須大於 0" },
            new() { Enabled = false, Field = "品名規格", RuleType = RuleType.MaxLength100,   Parameter = "", Message = "品名規格不得超過 100 字" },
        }
    };
}

public record ValidationError(int RowIndex, string OrderNo, string Field, string Message);
