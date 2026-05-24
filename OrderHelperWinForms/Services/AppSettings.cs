using System.Text.Json;
using OrderHelperWinForms.Models;

namespace OrderHelperWinForms.Services;

public static class AppSettings
{
    static readonly string ConfigDir = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
        "OrderHelper");

    static readonly JsonSerializerOptions JsonOpts = new()
    {
        WriteIndented = true,
        Converters = { new System.Text.Json.Serialization.JsonStringEnumConverter() },
    };

    // ---- Validation config ----

    static readonly string ValidationPath = Path.Combine(ConfigDir, "validation_config.json");

    public static ValidationConfig LoadValidation()
    {
        try
        {
            if (File.Exists(ValidationPath))
            {
                var json = File.ReadAllText(ValidationPath);
                return JsonSerializer.Deserialize<ValidationConfig>(json, JsonOpts)
                       ?? ValidationConfig.Default();
            }
        }
        catch { /* corrupt file — fall through to default */ }
        return ValidationConfig.Default();
    }

    public static void SaveValidation(ValidationConfig config)
    {
        EnsureDir();
        File.WriteAllText(ValidationPath, JsonSerializer.Serialize(config, JsonOpts));
    }

    // ---- Hospital settings ----

    static readonly string HospitalPath = Path.Combine(ConfigDir, "hospital_settings.json");

    public static HospitalSettings LoadHospital()
    {
        try
        {
            if (File.Exists(HospitalPath))
            {
                var json = File.ReadAllText(HospitalPath);
                return JsonSerializer.Deserialize<HospitalSettings>(json, JsonOpts)
                       ?? HospitalSettings.Default();
            }
        }
        catch { }
        return HospitalSettings.Default();
    }

    public static void SaveHospital(HospitalSettings settings)
    {
        EnsureDir();
        File.WriteAllText(HospitalPath, JsonSerializer.Serialize(settings, JsonOpts));
    }

    static void EnsureDir() => Directory.CreateDirectory(ConfigDir);
}
