using MaterialSkin;
using MaterialSkin.Controls;

namespace OutlookClassicSearch;

static class Program
{
    /// <summary>
    ///  The main entry point for the application.
    /// </summary>
    [STAThread]
    static void Main()
    {
        // To customize application configuration such as set high DPI settings or default font,
        // see https://aka.ms/applicationconfiguration.
        ApplicationConfiguration.Initialize();

        // Laad instellingen
        var settings = AppSettingsStore.Load();
        Strings.IsEnglish = settings.Language == "en";

        // Stel AppTheme in op basis van opgeslagen voorkeur
        AppTheme.CurrentMode = settings.Theme switch
        {
            "Dark" => AppTheme.ThemeMode.Dark,
            "Auto" => AppTheme.ThemeMode.Auto,
            _ => AppTheme.ThemeMode.Light
        };

        // MaterialSkin initialiseren
        var materialSkinManager = MaterialSkinManager.Instance;
        materialSkinManager.EnforceBackcolorOnAllComponents = false; // Uitgeschakeld om MenuStrip te laten werken
        materialSkinManager.Theme = AppTheme.IsDarkTheme 
            ? MaterialSkinManager.Themes.DARK 
            : MaterialSkinManager.Themes.LIGHT;
        materialSkinManager.ColorScheme = new ColorScheme(
            Primary.Blue600, 
            Primary.Blue700, 
            Primary.Blue200, 
            Accent.LightBlue200, 
            TextShade.WHITE
        );

        Application.Run(new Form1());
    }    
}