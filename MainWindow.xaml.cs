using System.IO;
using System.Text.Json;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using MonTableurApp.ViewModels;
using MonTableurApp.Views;

namespace MonTableurApp
{
    public partial class MainWindow : Window
    {
        private const string ThemeRose = "rose";
        private const string ThemeBlue = "blue";
        private const double SidebarExpandedWidth = 315;
        private const double SidebarCollapsedWidth = 144;
        private static readonly string ThemeSettingsPath = Path.Combine(AppContext.BaseDirectory, "ui-settings.json");

        private readonly MainViewModel viewModel = new MainViewModel();
        private readonly VueGeneraleView vueGenerale;
        private readonly VueSuiviEssaisView vueSuiviEssais;
        private readonly VueAjouterProjetView vueAjouterProjet;
        private readonly VueModifierProprietesView vueModifierProprietes;
        private readonly VueEnCoursView vueEnCours;
        private readonly VueAgendaView vueAgenda;
        private bool isBlueTheme;
        private bool isSidebarCollapsed;

        public MainWindow()
        {
            InitializeComponent();
            DataContext = viewModel;
            vueGenerale = new VueGeneraleView { DataContext = viewModel };
            vueSuiviEssais = new VueSuiviEssaisView { DataContext = viewModel };
            vueAjouterProjet = new VueAjouterProjetView { DataContext = viewModel };
            vueModifierProprietes = new VueModifierProprietesView { DataContext = viewModel };
            vueEnCours = new VueEnCoursView { DataContext = viewModel };
            vueAgenda = new VueAgendaView { DataContext = viewModel };

            isBlueTheme = LoadSavedTheme() == ThemeBlue;
            isSidebarCollapsed = LoadSavedSidebarCollapsed();
            ApplyCurrentTheme();
            ApplySidebarState();
            ApplySavedWindowPlacement();
            AfficherVueGenerale();
            Closing += MainWindow_Closing;
        }

        private void VueGenerale_Click(object sender, RoutedEventArgs e)
        {
            AfficherVueGenerale();
        }

        private void SuiviEssais_Click(object sender, RoutedEventArgs e)
        {
            AfficherVueSuiviEssais();
        }

        private void AjouterProjet_Click(object sender, RoutedEventArgs e)
        {
            AfficherVueAjouterProjet();
        }

        private void ModifierProprietes_Click(object sender, RoutedEventArgs e)
        {
            AfficherVueModifierProprietes();
        }

        private void EnCours_Click(object sender, RoutedEventArgs e)
        {
            AfficherVueEnCours();
        }

        private void Agenda_Click(object sender, RoutedEventArgs e)
        {
            AfficherVueAgenda();
        }

        private void ThemeToggle_Click(object sender, RoutedEventArgs e)
        {
            isBlueTheme = !isBlueTheme;
            ApplyCurrentTheme();
        }

        private void SidebarToggle_Click(object sender, RoutedEventArgs e)
        {
            isSidebarCollapsed = !isSidebarCollapsed;
            ApplySidebarState();
            SaveTheme();
        }

        private void MainWindow_Closing(object? sender, System.ComponentModel.CancelEventArgs e)
        {
            SaveTheme();
        }

        private void AfficherVueGenerale()
        {
            ShowView(
                vueGenerale,
                "Vue générale des projets",
                VueGeneraleButton);
        }

        private void AfficherVueSuiviEssais()
        {
            ShowView(
                vueSuiviEssais,
                "Suivi des essais",
                SuiviEssaisButton);
        }

        private void AfficherVueAjouterProjet()
        {
            ShowView(
                vueAjouterProjet,
                "Ajouter un projet",
                "Un espace prêt pour la future création de projets.",
                AjouterProjetButton);
        }

        private void AfficherVueModifierProprietes()
        {
            ShowView(
                vueModifierProprietes,
                "Modifier des propriétés",
                ModifierProprietesButton);
        }

        private void AfficherVueEnCours()
        {
            ShowView(
                vueEnCours,
                "En cours",
                "Cette vue accueillera bientôt les projets et essais en mouvement.",
                EnCoursButton);
        }

        private void AfficherVueAgenda()
        {
            ShowView(
                vueAgenda,
                "Agenda",
                "Une page vide pour préparer le planning et les échéances.",
                AgendaButton);
        }

        private void ShowView(UserControl view, string title, Button activeButton)
        {
            PageTitleText.Text = title;
            SetActiveButton(activeButton);
            MainContent.Content = view;
        }

        private void ShowView(UserControl view, string title, string subtitle, Button activeButton)
        {
            ShowView(view, title, activeButton);
        }

        private void SetActiveButton(Button activeButton)
        {
            VueGeneraleButton.Tag = null;
            SuiviEssaisButton.Tag = null;
            AjouterProjetButton.Tag = null;
            ModifierProprietesButton.Tag = null;
            EnCoursButton.Tag = null;
            AgendaButton.Tag = null;
            activeButton.Tag = "Active";
        }

        private void ApplySidebarState()
        {
            SidebarColumn.Width = new GridLength(isSidebarCollapsed ? SidebarCollapsedWidth : SidebarExpandedWidth);
            SidebarBorder.Margin = isSidebarCollapsed ? new Thickness(12) : new Thickness(18);
            SidebarInnerGrid.Margin = isSidebarCollapsed ? new Thickness(14) : new Thickness(28);
            NavigationPanel.Margin = isSidebarCollapsed ? new Thickness(0, 22, 0, 0) : new Thickness(0, 38, 0, 0);

            UserGreetingPanel.Visibility = isSidebarCollapsed ? Visibility.Collapsed : Visibility.Visible;
            AppTitleText.Visibility = isSidebarCollapsed ? Visibility.Collapsed : Visibility.Visible;

            LogoChip.Width = isSidebarCollapsed ? 54 : 84;
            LogoChip.Height = isSidebarCollapsed ? 54 : 84;
            LogoChip.CornerRadius = isSidebarCollapsed ? new CornerRadius(18) : new CornerRadius(26);
            LogoImage.Width = isSidebarCollapsed ? 42 : 64;
            LogoImage.Height = isSidebarCollapsed ? 42 : 64;

            SidebarToggleButton.Width = isSidebarCollapsed ? 34 : 42;
            SidebarToggleButton.Height = isSidebarCollapsed ? 32 : 36;
            SidebarToggleButton.Content = isSidebarCollapsed ? ">>" : "<<";
            SidebarToggleButton.ToolTip = isSidebarCollapsed
                ? "Afficher le bandeau de navigation"
                : "Réduire le bandeau de navigation";

            SetNavigationButtonLabel(VueGeneraleButton, "Vue générale", "V");
            SetNavigationButtonLabel(SuiviEssaisButton, "Suivi des essais", "S");
            SetNavigationButtonLabel(AjouterProjetButton, "Ajouter un projet", "+");
            SetNavigationButtonLabel(EnCoursButton, "En cours", "C");
            SetNavigationButtonLabel(AgendaButton, "Agenda", "A");
            SetNavigationButtonLabel(ModifierProprietesButton, "Modifier des propriétés", "M");
        }

        private void SetNavigationButtonLabel(Button button, string fullLabel, string compactLabel)
        {
            button.Content = isSidebarCollapsed ? compactLabel : fullLabel;
            button.ToolTip = isSidebarCollapsed ? fullLabel : null;
            button.HorizontalContentAlignment = isSidebarCollapsed ? HorizontalAlignment.Center : HorizontalAlignment.Left;
            button.Padding = isSidebarCollapsed ? new Thickness(6, 12, 6, 14) : new Thickness(10, 12, 10, 14);
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
            ThemeSettings? settings = LoadThemeSettings();
            return settings?.Theme == ThemeBlue ? ThemeBlue : ThemeRose;
        }

        private static bool LoadSavedSidebarCollapsed()
        {
            ThemeSettings? settings = LoadThemeSettings();
            return settings?.SidebarCollapsed == true;
        }

        private static ThemeSettings? LoadThemeSettings()
        {
            if (!File.Exists(ThemeSettingsPath))
            {
                return null;
            }

            try
            {
                string json = File.ReadAllText(ThemeSettingsPath);
                return JsonSerializer.Deserialize<ThemeSettings>(json);
            }
            catch
            {
                return null;
            }
        }

        private void ApplySavedWindowPlacement()
        {
            ThemeSettings? settings = LoadThemeSettings();
            if (settings == null)
            {
                return;
            }

            if (IsValidWindowSize(settings.WindowWidth, settings.WindowHeight))
            {
                Width = Math.Max(MinWidth, settings.WindowWidth!.Value);
                Height = Math.Max(MinHeight, settings.WindowHeight!.Value);

                if (IsValidWindowPosition(settings.WindowLeft, settings.WindowTop, Width, Height))
                {
                    WindowStartupLocation = WindowStartupLocation.Manual;
                    Left = settings.WindowLeft!.Value;
                    Top = settings.WindowTop!.Value;
                }
            }

            if (Enum.TryParse(settings.WindowState, out WindowState savedState) &&
                savedState != WindowState.Minimized)
            {
                WindowState = savedState;
            }
        }

        private static bool IsValidWindowSize(double? width, double? height)
        {
            return width.HasValue &&
                   height.HasValue &&
                   double.IsFinite(width.Value) &&
                   double.IsFinite(height.Value) &&
                   width.Value > 0 &&
                   height.Value > 0;
        }

        private static bool IsValidWindowPosition(double? left, double? top, double width, double height)
        {
            if (!left.HasValue ||
                !top.HasValue ||
                !double.IsFinite(left.Value) ||
                !double.IsFinite(top.Value))
            {
                return false;
            }

            var savedBounds = new Rect(left.Value, top.Value, width, height);
            var virtualScreen = new Rect(
                SystemParameters.VirtualScreenLeft,
                SystemParameters.VirtualScreenTop,
                SystemParameters.VirtualScreenWidth,
                SystemParameters.VirtualScreenHeight);

            return savedBounds.IntersectsWith(virtualScreen);
        }

        private void SaveTheme()
        {
            Rect normalBounds = WindowState == WindowState.Normal
                ? new Rect(Left, Top, Width, Height)
                : RestoreBounds;

            var settings = new ThemeSettings
            {
                Theme = isBlueTheme ? ThemeBlue : ThemeRose,
                SidebarCollapsed = isSidebarCollapsed,
                WindowLeft = normalBounds.Left,
                WindowTop = normalBounds.Top,
                WindowWidth = normalBounds.Width,
                WindowHeight = normalBounds.Height,
                WindowState = WindowState == WindowState.Minimized
                    ? nameof(WindowState.Normal)
                    : WindowState.ToString()
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
            resources["DataGridRowSelectedBrush"] = CreateBrush("#FFEAF3");
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

            resources["WindowBackgroundBrush"] = CreateGradient("#F2F8FF", "#E7F1FF", "#F4F1FF");
            resources["SidebarBrush"] = CreateGradient("#CFE4FF", "#D4D8FF", "#E7DBFF");
            resources["SidebarBorderBrush"] = CreateBrush("#9CBDF0");
            resources["LogoChipBackgroundBrush"] = CreateBrush("#FDFEFF");
            resources["LogoChipBorderBrush"] = CreateBrush("#A8C8F3");
            resources["PrimaryTitleBrush"] = CreateBrush("#3F5FA6");
            resources["SecondaryTextBrush"] = CreateBrush("#657FB3");

            resources["MenuButtonBackgroundBrush"] = CreateBrush("#FCFEFF");
            resources["MenuButtonBorderBrush"] = CreateBrush("#B9D1F7");
            resources["MenuButtonForegroundBrush"] = CreateBrush("#4865A7");
            resources["MenuButtonHoverBackgroundBrush"] = CreateBrush("#E6F0FF");
            resources["MenuButtonHoverBorderBrush"] = CreateBrush("#7EA5E8");
            resources["MenuButtonPressedBackgroundBrush"] = CreateBrush("#D6E7FF");

            resources["VersionBadgeBackgroundBrush"] = CreateBrush("#EAF1FF");
            resources["VersionBadgeBorderBrush"] = CreateBrush("#ABC5EE");
            resources["VersionBadgeForegroundBrush"] = CreateBrush("#5474BC");

            resources["ThemeButtonBackgroundBrush"] = CreateBrush("#FCFEFF");
            resources["ThemeButtonBorderBrush"] = CreateBrush("#BFD2F5");
            resources["ThemeButtonForegroundBrush"] = CreateBrush("#4763A5");

            resources["InfoCardBackgroundBrush"] = CreateBrush("#F2F7FF");
            resources["InfoCardBorderBrush"] = CreateBrush("#B9D0F3");
            resources["InfoCardAccentBrush"] = CreateBrush("#5F8DE5");

            resources["ContentBackgroundBrush"] = CreateBrush("#FCFEFF");
            resources["ContentBorderBrush"] = CreateBrush("#C7D9F5");

            resources["SummaryCardBackgroundBrush"] = CreateBrush("#EEF5FF");
            resources["SummaryCardBorderBrush"] = CreateBrush("#BED1F1");
            resources["SummaryLabelBrush"] = CreateBrush("#6F84B2");
            resources["SummaryValueBrush"] = CreateBrush("#4265AF");

            resources["SummaryWarmBackgroundBrush"] = CreateBrush("#E5F2FF");
            resources["SummaryWarmBorderBrush"] = CreateBrush("#AED0F6");
            resources["SummaryWarmForegroundBrush"] = CreateBrush("#4470BA");

            resources["SummaryMintBackgroundBrush"] = CreateBrush("#E5FAFF");
            resources["SummaryMintBorderBrush"] = CreateBrush("#A7DDF2");
            resources["SummaryMintForegroundBrush"] = CreateBrush("#2F86A6");

            resources["SummaryLavenderBackgroundBrush"] = CreateBrush("#ECEBFF");
            resources["SummaryLavenderBorderBrush"] = CreateBrush("#BFC4F7");
            resources["SummaryLavenderForegroundBrush"] = CreateBrush("#5864C8");

            resources["SearchCardBackgroundBrush"] = CreateBrush("#F5F9FF");
            resources["SearchCardBorderBrush"] = CreateBrush("#BED1F1");
            resources["SearchInputBackgroundBrush"] = CreateBrush("#FFFFFF");
            resources["SearchInputBorderBrush"] = CreateBrush("#B7CEF3");
            resources["SearchInputForegroundBrush"] = CreateBrush("#506C9F");

            resources["DataGridHeaderBrush"] = CreateBrush("#DAEBFF");
            resources["DataGridHeaderForegroundBrush"] = CreateBrush("#4666AA");
            resources["DataGridRowHoverBrush"] = CreateBrush("#EFF5FF");
            resources["DataGridRowSelectedBrush"] = CreateBrush("#DCEBFF");
            resources["DataGridSelectionForegroundBrush"] = CreateBrush("#3D5793");
            resources["DataGridAltRowBrush"] = CreateBrush("#F7FAFF");
            resources["DataGridSurfaceBorderBrush"] = CreateBrush("#C8D8F0");

            resources["ScrollTrackBrush"] = CreateBrush("#D6E3F8");
            resources["ScrollThumbBrush"] = CreateBrush("#7EA5E8");
            resources["ScrollThumbHoverBrush"] = CreateBrush("#5E88D7");
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

            public bool SidebarCollapsed { get; set; }

            public double? WindowLeft { get; set; }

            public double? WindowTop { get; set; }

            public double? WindowWidth { get; set; }

            public double? WindowHeight { get; set; }

            public string? WindowState { get; set; }
        }
    }
}
