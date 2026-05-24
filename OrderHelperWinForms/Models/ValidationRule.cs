using System.Text.Json.Serialization;

namespace OrderHelperWinForms.Models;

public enum RuleType { Required, Regex, MaxLength }

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
            new() { Enabled = true,  Field = "訂購量",   RuleType = RuleType.Regex,     Parameter = @"^\d+(\.\d+)?$", Message = "訂購量必須為數字" },
            new() { Enabled = false, Field = "品名規格", RuleType = RuleType.MaxLength, Parameter = "100",           Message = "品名規格不得超過100字" },
        }
    };
}

public record ValidationError(int RowIndex, string OrderNo, string Field, string Message);
