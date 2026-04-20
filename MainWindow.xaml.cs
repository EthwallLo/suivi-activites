using System.IO;
using System.Text.Json;
using System.Windows;
using System.Windows.Media;
using MonTableurApp.ViewModels;
using MonTableurApp.Views;

namespace MonTableurApp
{
    public partial class MainWindow : Window
    {
        private const string ThemeRose = "rose";
        private const string ThemeBlue = "blue";
        private static readonly string ThemeSettingsPath = Path.Combine(AppContext.BaseDirectory, "ui-settings.json");

        private readonly MainViewModel viewModel = new MainViewModel();
        private bool isBlueTheme;

        public MainWindow()
        {
            InitializeComponent();
            DataContext = viewModel;

            isBlueTheme = LoadSavedTheme() == ThemeBlue;
            ApplyCurrentTheme();
            AfficherVueGenerale();
            Closing += MainWindow_Closing;
        }

        private void VueGenerale_Click(object sender, RoutedEventArgs e)
        {
            AfficherVueGenerale();
        }

        private void ThemeToggle_Click(object sender, RoutedEventArgs e)
        {
            isBlueTheme = !isBlueTheme;
            ApplyCurrentTheme();
        }

        private void MainWindow_Closing(object? sender, System.ComponentModel.CancelEventArgs e)
        {
            SaveTheme();
        }

        private void AfficherVueGenerale()
        {
            var vue = new VueGeneraleView
            {
                DataContext = viewModel
            };

            MainContent.Content = vue;
        }

        private void ApplyCurrentTheme()
        {
            if (isBlueTheme)
            {
                ApplyBlueTheme();
                ThemeToggleButton.Content = "Rose";
            }
            else
            {
                ApplyRoseTheme();
                ThemeToggleButton.Content = "Bleu";
            }
        }

        private static string LoadSavedTheme()
        {
            if (!File.Exists(ThemeSettingsPath))
            {
                return ThemeRose;
            }

            try
            {
                string json = File.ReadAllText(ThemeSettingsPath);
                ThemeSettings? settings = JsonSerializer.Deserialize<ThemeSettings>(json);
                return settings?.Theme == ThemeBlue ? ThemeBlue : ThemeRose;
            }
            catch
            {
                return ThemeRose;
            }
        }

        private void SaveTheme()
        {
            var settings = new ThemeSettings
            {
                Theme = isBlueTheme ? ThemeBlue : ThemeRose
            };

            string json = JsonSerializer.Serialize(settings, new JsonSerializerOptions { WriteIndented = true });
            File.WriteAllText(ThemeSettingsPath, json);
        }

        private static void ApplyRoseTheme()
        {
            ResourceDictionary resources = Application.Current.Resources;

            resources["WindowBackgroundBrush"] = CreateGradient("#FFF7FC", "#FFF0F7", "#FFF8EF");
            resources["SidebarBrush"] = CreateGradient("#FFD7E8", "#F6D8FF");
            resources["SidebarBorderBrush"] = CreateBrush("#F5B7D7");
            resources["LogoChipBackgroundBrush"] = CreateBrush("#FFF9FD");
            resources["LogoChipBorderBrush"] = CreateBrush("#F4B5D6");
            resources["PrimaryTitleBrush"] = CreateBrush("#8B4B7C");
            resources["SecondaryTextBrush"] = CreateBrush("#9B758A");

            resources["MenuButtonBackgroundBrush"] = CreateBrush("#FFFDFE");
            resources["MenuButtonBorderBrush"] = CreateBrush("#F6BAD8");
            resources["MenuButtonForegroundBrush"] = CreateBrush("#7B4D73");
            resources["MenuButtonHoverBackgroundBrush"] = CreateBrush("#FFF0F8");
            resources["MenuButtonHoverBorderBrush"] = CreateBrush("#EFA6CC");
            resources["MenuButtonPressedBackgroundBrush"] = CreateBrush("#FFE7F4");

            resources["VersionBadgeBackgroundBrush"] = CreateBrush("#FFF4D9");
            resources["VersionBadgeBorderBrush"] = CreateBrush("#F5D59A");
            resources["VersionBadgeForegroundBrush"] = CreateBrush("#B78641");

            resources["ThemeButtonBackgroundBrush"] = CreateBrush("#FFFDFE");
            resources["ThemeButtonBorderBrush"] = CreateBrush("#E7CEE4");
            resources["ThemeButtonForegroundBrush"] = CreateBrush("#7B4D73");

            resources["InfoCardBackgroundBrush"] = CreateBrush("#FFF6FB");
            resources["InfoCardBorderBrush"] = CreateBrush("#F5BCD9");
            resources["InfoCardAccentBrush"] = CreateBrush("#E28AB9");

            resources["ContentBackgroundBrush"] = CreateBrush("#FFFCFE");
            resources["ContentBorderBrush"] = CreateBrush("#F4D4E5");

            resources["SummaryCardBackgroundBrush"] = CreateBrush("#FFF2F8");
            resources["SummaryCardBorderBrush"] = CreateBrush("#F5C8DD");
            resources["SummaryLabelBrush"] = CreateBrush("#A07A94");
            resources["SummaryValueBrush"] = CreateBrush("#8A4E78");

            resources["SummaryWarmBackgroundBrush"] = CreateBrush("#FFF4D8");
            resources["SummaryWarmBorderBrush"] = CreateBrush("#F1D89B");
            resources["SummaryWarmForegroundBrush"] = CreateBrush("#A8712A");

            resources["SummaryMintBackgroundBrush"] = CreateBrush("#EAF8F5");
            resources["SummaryMintBorderBrush"] = CreateBrush("#B8E4D9");
            resources["SummaryMintForegroundBrush"] = CreateBrush("#3D897A");

            resources["SummaryLavenderBackgroundBrush"] = CreateBrush("#F0EEFF");
            resources["SummaryLavenderBorderBrush"] = CreateBrush("#D3CEF8");
            resources["SummaryLavenderForegroundBrush"] = CreateBrush("#625AB1");

            resources["SearchCardBackgroundBrush"] = CreateBrush("#FFF9FC");
            resources["SearchCardBorderBrush"] = CreateBrush("#F5C8DD");
            resources["SearchInputBackgroundBrush"] = CreateBrush("#FFFFFF");
            resources["SearchInputBorderBrush"] = CreateBrush("#EFC3DA");
            resources["SearchInputForegroundBrush"] = CreateBrush("#7A5B72");

            resources["DataGridHeaderBrush"] = CreateBrush("#FFE6F1");
            resources["DataGridHeaderForegroundBrush"] = CreateBrush("#8D4D79");
            resources["DataGridRowHoverBrush"] = CreateBrush("#FFF5FA");
            resources["DataGridRowSelectedBrush"] = CreateBrush("#FFE8F4");
            resources["DataGridSelectionForegroundBrush"] = CreateBrush("#6E4967");
            resources["DataGridAltRowBrush"] = CreateBrush("#FFF9FC");
            resources["DataGridSurfaceBorderBrush"] = CreateBrush("#F3D9E7");

            resources["ScrollTrackBrush"] = CreateBrush("#F7DDEB");
            resources["ScrollThumbBrush"] = CreateBrush("#E9A8CC");
            resources["ScrollThumbHoverBrush"] = CreateBrush("#D98FB8");
        }

        private static void ApplyBlueTheme()
        {
            ResourceDictionary resources = Application.Current.Resources;

            resources["WindowBackgroundBrush"] = CreateGradient("#F5FBFF", "#EEF6FF", "#F5F3FF");
            resources["SidebarBrush"] = CreateGradient("#D9ECFF", "#DDE5FF");
            resources["SidebarBorderBrush"] = CreateBrush("#B8D4F5");
            resources["LogoChipBackgroundBrush"] = CreateBrush("#FDFEFF");
            resources["LogoChipBorderBrush"] = CreateBrush("#BBD6F7");
            resources["PrimaryTitleBrush"] = CreateBrush("#4B6E9E");
            resources["SecondaryTextBrush"] = CreateBrush("#6D83A7");

            resources["MenuButtonBackgroundBrush"] = CreateBrush("#FCFEFF");
            resources["MenuButtonBorderBrush"] = CreateBrush("#C5DAF8");
            resources["MenuButtonForegroundBrush"] = CreateBrush("#4E6C97");
            resources["MenuButtonHoverBackgroundBrush"] = CreateBrush("#EEF7FF");
            resources["MenuButtonHoverBorderBrush"] = CreateBrush("#AACAF0");
            resources["MenuButtonPressedBackgroundBrush"] = CreateBrush("#E1F0FF");

            resources["VersionBadgeBackgroundBrush"] = CreateBrush("#EAF4FF");
            resources["VersionBadgeBorderBrush"] = CreateBrush("#BDD6F3");
            resources["VersionBadgeForegroundBrush"] = CreateBrush("#5B82B8");

            resources["ThemeButtonBackgroundBrush"] = CreateBrush("#FCFEFF");
            resources["ThemeButtonBorderBrush"] = CreateBrush("#CDDDF6");
            resources["ThemeButtonForegroundBrush"] = CreateBrush("#4E6C97");

            resources["InfoCardBackgroundBrush"] = CreateBrush("#F7FBFF");
            resources["InfoCardBorderBrush"] = CreateBrush("#C8DBF4");
            resources["InfoCardAccentBrush"] = CreateBrush("#78A5DD");

            resources["ContentBackgroundBrush"] = CreateBrush("#FCFEFF");
            resources["ContentBorderBrush"] = CreateBrush("#D5E3F8");

            resources["SummaryCardBackgroundBrush"] = CreateBrush("#F2F8FF");
            resources["SummaryCardBorderBrush"] = CreateBrush("#C9DCF5");
            resources["SummaryLabelBrush"] = CreateBrush("#7C8FAF");
            resources["SummaryValueBrush"] = CreateBrush("#5574A4");

            resources["SummaryWarmBackgroundBrush"] = CreateBrush("#EDF5FF");
            resources["SummaryWarmBorderBrush"] = CreateBrush("#C5DCF6");
            resources["SummaryWarmForegroundBrush"] = CreateBrush("#5A80B0");

            resources["SummaryMintBackgroundBrush"] = CreateBrush("#EAFBFF");
            resources["SummaryMintBorderBrush"] = CreateBrush("#C3EAF4");
            resources["SummaryMintForegroundBrush"] = CreateBrush("#4A8CA5");

            resources["SummaryLavenderBackgroundBrush"] = CreateBrush("#EEF0FF");
            resources["SummaryLavenderBorderBrush"] = CreateBrush("#CCD3FA");
            resources["SummaryLavenderForegroundBrush"] = CreateBrush("#5D6EC4");

            resources["SearchCardBackgroundBrush"] = CreateBrush("#F7FBFF");
            resources["SearchCardBorderBrush"] = CreateBrush("#C8DBF4");
            resources["SearchInputBackgroundBrush"] = CreateBrush("#FFFFFF");
            resources["SearchInputBorderBrush"] = CreateBrush("#C7DBF6");
            resources["SearchInputForegroundBrush"] = CreateBrush("#5D7499");

            resources["DataGridHeaderBrush"] = CreateBrush("#E8F3FF");
            resources["DataGridHeaderForegroundBrush"] = CreateBrush("#5877A8");
            resources["DataGridRowHoverBrush"] = CreateBrush("#F3F8FF");
            resources["DataGridRowSelectedBrush"] = CreateBrush("#E4F0FF");
            resources["DataGridSelectionForegroundBrush"] = CreateBrush("#4B638B");
            resources["DataGridAltRowBrush"] = CreateBrush("#F9FBFF");
            resources["DataGridSurfaceBorderBrush"] = CreateBrush("#D8E5F8");

            resources["ScrollTrackBrush"] = CreateBrush("#DDEAF8");
            resources["ScrollThumbBrush"] = CreateBrush("#9FC2EB");
            resources["ScrollThumbHoverBrush"] = CreateBrush("#84AFDF");
        }

        private static SolidColorBrush CreateBrush(string color)
        {
            var brush = (SolidColorBrush)new BrushConverter().ConvertFromString(color)!;
            brush.Freeze();
            return brush;
        }

        private static LinearGradientBrush CreateGradient(string startColor, string endColor)
        {
            return CreateGradient(startColor, endColor, endColor);
        }

        private static LinearGradientBrush CreateGradient(string startColor, string middleColor, string endColor)
        {
            var brush = new LinearGradientBrush
            {
                StartPoint = new Point(0, 0),
                EndPoint = new Point(1, 1)
            };

            brush.GradientStops.Add(new GradientStop((Color)ColorConverter.ConvertFromString(startColor)!, 0));
            brush.GradientStops.Add(new GradientStop((Color)ColorConverter.ConvertFromString(middleColor)!, 0.55));
            brush.GradientStops.Add(new GradientStop((Color)ColorConverter.ConvertFromString(endColor)!, 1));
            brush.Freeze();

            return brush;
        }

        private sealed class ThemeSettings
        {
            public string Theme { get; set; } = ThemeRose;
        }
    }
}
