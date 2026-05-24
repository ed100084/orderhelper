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
        catch (JsonException) { }
        catch (IOException) { }
        catch (UnauthorizedAccessException) { }
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
        catch (JsonException) { }
        catch (IOException) { }
        catch (UnauthorizedAccessException) { }
        return HospitalSettings.Default();
    }

    public static void SaveHospital(HospitalSettings settings)
    {
        EnsureDir();
        File.WriteAllText(HospitalPath, JsonSerializer.Serialize(settings, JsonOpts));
    }

    // ---- General settings ----

    static readonly string GeneralPath = Path.Combine(ConfigDir, "general_settings.json");

    public static GeneralSettings LoadGeneral()
    {
        try
        {
            if (File.Exists(GeneralPath))
            {
                var json = File.ReadAllText(GeneralPath);
                return JsonSerializer.Deserialize<GeneralSettings>(json, JsonOpts)
                       ?? GeneralSettings.Default();
            }
        }
        catch (JsonException) { }
        catch (IOException) { }
        catch (UnauthorizedAccessException) { }
        return GeneralSettings.Default();
    }

    public static void SaveGeneral(GeneralSettings settings)
    {
        EnsureDir();
        File.WriteAllText(GeneralPath, JsonSerializer.Serialize(settings, JsonOpts));
    }

    static void EnsureDir() => Directory.CreateDirectory(ConfigDir);
}
