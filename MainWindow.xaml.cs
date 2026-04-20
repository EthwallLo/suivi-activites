using System.Windows;
using MonTableurApp.ViewModels;
using MonTableurApp.Views;

namespace MonTableurApp
{
    public partial class MainWindow : Window
    {
        private readonly MainViewModel viewModel = new MainViewModel();

        public MainWindow()
        {
            InitializeComponent();
            DataContext = viewModel;

            AfficherVueGenerale();
        }

        private void VueGenerale_Click(object sender, RoutedEventArgs e)
        {
            AfficherVueGenerale();
        }

        private void AfficherVueGenerale()
        {
            var vue = new VueGeneraleView
            {
                DataContext = viewModel
            };

            MainContent.Content = vue;
        }
    }
}
