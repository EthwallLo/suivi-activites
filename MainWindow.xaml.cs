using System.Windows;
using MonTableurApp.ViewModels;
using MonTableurApp.Views;

namespace MonTableurApp
{
    public partial class MainWindow : Window
    {
        private MainViewModel viewModel = new MainViewModel();

        public MainWindow()
        {
            InitializeComponent();

            var vue = new VueGeneraleView();
            vue.DataContext = viewModel;
            MainContent.Content = vue;
        }

        private void VueGenerale_Click(object sender, RoutedEventArgs e)
        {
            var vue = new VueGeneraleView();
            vue.DataContext = viewModel;
            MainContent.Content = vue;
        }
    }
}