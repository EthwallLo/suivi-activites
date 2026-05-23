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
        private readonly VueArchivesView vueArchives;
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
            vueArchives = new VueArchivesView { DataContext = viewModel };

            isSidebarCollapsed = LoadSavedSidebarCollapsed();
            ApplyApplicationTheme();
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

        private void Archives_Click(object sender, RoutedEventArgs e)
        {
            AfficherVueArchives();
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

        private void AfficherVueArchives()
        {
            ShowView(
                vueArchives,
                "Afficher les projets archivés",
                ArchivesButton);
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
            ArchivesButton.Tag = null;
            activeButton.Tag = "Active";
        }

        private void ApplySidebarState()
        {
            SidebarColumn.Width = new GridLength(isSidebarCollapsed ? SidebarCollapsedWidth : SidebarExpandedWidth);
            SidebarBorder.Margin = isSidebarCollapsed ? new Thickness(12) : new Thickness(18);
            SidebarInnerGrid.Margin = isSidebarCollapsed ? new Thickness(14) : new Thickness(28);
            NavigationPanel.Margin = isSidebarCollapsed ? new Thickness(0, 14, 0, 0) : new Thickness(0, 18, 0, 0);
            AppTitleText.Visibility = isSidebarCollapsed ? Visibility.Collapsed : Visibility.Visible;

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
            SetNavigationButtonLabel(ArchivesButton, "Projets archivés", "P");
            SetNavigationButtonLabel(ModifierProprietesButton, "Modifier des propriétés", "M");
        }

        private void SetNavigationButtonLabel(Button button, string fullLabel, string compactLabel)
        {
            button.Content = isSidebarCollapsed ? compactLabel : fullLabel;
            button.ToolTip = isSidebarCollapsed ? fullLabel : null;
            button.HorizontalContentAlignment = isSidebarCollapsed ? HorizontalAlignment.Center : HorizontalAlignment.Left;
            button.Padding = isSidebarCollapsed ? new Thickness(6, 12, 6, 14) : new Thickness(10, 12, 10, 14);
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

        private static void ApplyApplicationTheme()
        {
            ResourceDictionary resources = Application.Current.Resources;

            resources["WindowBackgroundBrush"] = CreateGradient("#F4F8FC", "#EEF3F9", "#E6EDF6");
            resources["SidebarBrush"] = CreateGradient("#E8F0F8", "#DBE7F2");
            resources["SidebarBorderBrush"] = CreateBrush("#C3D3E3");
            resources["LogoChipBackgroundBrush"] = CreateBrush("#FFFFFF");
            resources["LogoChipBorderBrush"] = CreateBrush("#D0DCE9");
            resources["PrimaryTitleBrush"] = CreateBrush("#1F3B5B");
            resources["SecondaryTextBrush"] = CreateBrush("#5E7590");

            resources["MenuButtonBackgroundBrush"] = CreateBrush("#FFFFFF");
            resources["MenuButtonBorderBrush"] = CreateBrush("#C5D6E6");
            resources["MenuButtonForegroundBrush"] = CreateBrush("#274969");
            resources["MenuButtonHoverBackgroundBrush"] = CreateBrush("#E7F0F9");
            resources["MenuButtonHoverBorderBrush"] = CreateBrush("#6E92B7");
            resources["MenuButtonPressedBackgroundBrush"] = CreateBrush("#D9E6F3");

            resources["VersionBadgeBackgroundBrush"] = CreateBrush("#E9F0F7");
            resources["VersionBadgeBorderBrush"] = CreateBrush("#C6D4E2");
            resources["VersionBadgeForegroundBrush"] = CreateBrush("#31516F");

            resources["ThemeButtonBackgroundBrush"] = CreateBrush("#FFFFFF");
            resources["ThemeButtonBorderBrush"] = CreateBrush("#C8D5E3");
            resources["ThemeButtonForegroundBrush"] = CreateBrush("#274969");

            resources["InfoCardBackgroundBrush"] = CreateBrush("#EAF2FA");
            resources["InfoCardBorderBrush"] = CreateBrush("#C5D6E7");
            resources["InfoCardAccentBrush"] = CreateBrush("#3E709E");

            resources["ContentBackgroundBrush"] = CreateBrush("#FDFEFF");
            resources["ContentBorderBrush"] = CreateBrush("#D0DCE8");

            resources["SummaryCardBackgroundBrush"] = CreateBrush("#F2F7FC");
            resources["SummaryCardBorderBrush"] = CreateBrush("#CBD9E7");
            resources["SummaryLabelBrush"] = CreateBrush("#60768F");
            resources["SummaryValueBrush"] = CreateBrush("#2B4C6C");

            resources["SummaryWarmBackgroundBrush"] = CreateBrush("#EDF3F9");
            resources["SummaryWarmBorderBrush"] = CreateBrush("#CBD8E5");
            resources["SummaryWarmForegroundBrush"] = CreateBrush("#426483");

            resources["SummaryMintBackgroundBrush"] = CreateBrush("#EAF5F7");
            resources["SummaryMintBorderBrush"] = CreateBrush("#C2D8DE");
            resources["SummaryMintForegroundBrush"] = CreateBrush("#2F6D7B");

            resources["SummaryLavenderBackgroundBrush"] = CreateBrush("#EEF2F9");
            resources["SummaryLavenderBorderBrush"] = CreateBrush("#CCD5E6");
            resources["SummaryLavenderForegroundBrush"] = CreateBrush("#4E6683");

            resources["SearchCardBackgroundBrush"] = CreateBrush("#F6FAFD");
            resources["SearchCardBorderBrush"] = CreateBrush("#CFDBE7");
            resources["SearchInputBackgroundBrush"] = CreateBrush("#FFFFFF");
            resources["SearchInputBorderBrush"] = CreateBrush("#C4D4E4");
            resources["SearchInputForegroundBrush"] = CreateBrush("#2F4E6E");

            resources["DataGridHeaderBrush"] = CreateBrush("#E4EDF7");
            resources["DataGridHeaderForegroundBrush"] = CreateBrush("#2D4D6E");
            resources["DataGridRowHoverBrush"] = CreateBrush("#EDF4FB");
            resources["DataGridRowSelectedBrush"] = CreateBrush("#DDE9F7");
            resources["DataGridSelectionForegroundBrush"] = CreateBrush("#264562");
            resources["DataGridAltRowBrush"] = CreateBrush("#F9FBFE");
            resources["DataGridSurfaceBorderBrush"] = CreateBrush("#CCD9E6");

            resources["ScrollTrackBrush"] = CreateBrush("#D4E0EB");
            resources["ScrollThumbBrush"] = CreateBrush("#7898B8");
            resources["ScrollThumbHoverBrush"] = CreateBrush("#5C7E9F");
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
            public bool SidebarCollapsed { get; set; }

            public double? WindowLeft { get; set; }

            public double? WindowTop { get; set; }

            public double? WindowWidth { get; set; }

            public double? WindowHeight { get; set; }

            public string? WindowState { get; set; }
        }
    }
}
