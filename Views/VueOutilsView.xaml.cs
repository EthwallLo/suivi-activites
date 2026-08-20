using System.Globalization;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using Microsoft.Win32;

namespace MonTableurApp.Views
{
    public partial class VueOutilsView : UserControl
    {
        private const int MaxGeneratedNumberCount = 10_000;
        private readonly List<int> generatedNumbers = new List<int>();

        public VueOutilsView()
        {
            InitializeComponent();
        }

        private void Generate_Click(object sender, RoutedEventArgs e)
        {
            GenerateNumbers();
        }

        private void ParameterTextBox_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key != Key.Enter)
            {
                return;
            }

            GenerateNumbers();
            e.Handled = true;
        }

        private void GenerateNumbers()
        {
            ClearResults();

            if (!TryReadParameters(out int numberCount, out int minimum, out int maximum))
            {
                return;
            }

            generatedNumbers.Clear();
            long exclusiveMaximum = (long)maximum + 1;

            for (int index = 0; index < numberCount; index++)
            {
                long value = Random.Shared.NextInt64(minimum, exclusiveMaximum);
                generatedNumbers.Add((int)value);
            }

            ResultsTextBox.Text = string.Join(
                Environment.NewLine,
                generatedNumbers.Select(number => number.ToString(CultureInfo.InvariantCulture)));

            string noun = numberCount > 1 ? "nombres générés" : "nombre généré";
            ResultSummaryTextBlock.Text = $"{numberCount:N0} {noun} dans [{minimum:N0} ; {maximum:N0}]";
            CopyButton.IsEnabled = true;
            ExportCsvButton.IsEnabled = true;
            HideMessages();
        }

        private bool TryReadParameters(out int numberCount, out int minimum, out int maximum)
        {
            numberCount = 0;
            minimum = 0;
            maximum = 0;

            if (!TryParseInteger(NumberCountTextBox.Text, out numberCount))
            {
                ShowError("Le nombre de valeurs doit être un entier.");
                NumberCountTextBox.Focus();
                return false;
            }

            if (numberCount < 1 || numberCount > MaxGeneratedNumberCount)
            {
                ShowError($"Le nombre de valeurs doit être compris entre 1 et {MaxGeneratedNumberCount:N0}.");
                NumberCountTextBox.Focus();
                return false;
            }

            if (!TryParseInteger(MinimumTextBox.Text, out minimum))
            {
                ShowError("La valeur minimale doit être un entier compris entre -2 147 483 648 et 2 147 483 647.");
                MinimumTextBox.Focus();
                return false;
            }

            if (!TryParseInteger(MaximumTextBox.Text, out maximum))
            {
                ShowError("La valeur maximale doit être un entier compris entre -2 147 483 648 et 2 147 483 647.");
                MaximumTextBox.Focus();
                return false;
            }

            if (minimum > maximum)
            {
                ShowError("La valeur minimale doit être inférieure ou égale à la valeur maximale.");
                MinimumTextBox.Focus();
                return false;
            }

            return true;
        }

        private static bool TryParseInteger(string text, out int value)
        {
            string normalizedText = text
                .Replace(" ", string.Empty)
                .Replace("\u00A0", string.Empty)
                .Replace("\u202F", string.Empty);

            return int.TryParse(
                normalizedText,
                NumberStyles.Integer,
                CultureInfo.CurrentCulture,
                out value);
        }

        private void Copy_Click(object sender, RoutedEventArgs e)
        {
            if (generatedNumbers.Count == 0)
            {
                return;
            }

            try
            {
                Clipboard.SetText(ResultsTextBox.Text);
                string noun = generatedNumbers.Count > 1 ? "nombres copiés" : "nombre copié";
                ShowStatus($"{generatedNumbers.Count:N0} {noun} dans le presse-papiers.");
            }
            catch (ExternalException)
            {
                ShowError("Le presse-papiers est momentanément indisponible. Réessayez dans quelques secondes.");
            }
        }

        private void ExportCsv_Click(object sender, RoutedEventArgs e)
        {
            if (generatedNumbers.Count == 0)
            {
                return;
            }

            var dialog = new SaveFileDialog
            {
                Title = "Télécharger les nombres aléatoires",
                Filter = "Fichier CSV (*.csv)|*.csv",
                DefaultExt = ".csv",
                AddExtension = true,
                OverwritePrompt = true,
                FileName = $"nombres-aleatoires-{DateTime.Now:yyyy-MM-dd-HHmmss}.csv"
            };

            Window? owner = Window.GetWindow(this);
            bool? result = owner is null ? dialog.ShowDialog() : dialog.ShowDialog(owner);
            if (result != true)
            {
                return;
            }

            try
            {
                File.WriteAllText(dialog.FileName, BuildCsv(), new UTF8Encoding(encoderShouldEmitUTF8Identifier: true));
                string noun = generatedNumbers.Count > 1 ? "nombres exportés" : "nombre exporté";
                ShowStatus($"{generatedNumbers.Count:N0} {noun} dans le fichier CSV.");
            }
            catch (IOException)
            {
                ShowError("Le fichier CSV n'a pas pu être écrit. Vérifiez qu'il n'est pas déjà ouvert, puis réessayez.");
            }
            catch (UnauthorizedAccessException)
            {
                ShowError("Le fichier CSV n'a pas pu être écrit à cet emplacement. Choisissez un autre dossier.");
            }
        }

        private string BuildCsv()
        {
            var builder = new StringBuilder();
            builder.AppendLine("Nombre");

            foreach (int number in generatedNumbers)
            {
                builder.AppendLine(number.ToString(CultureInfo.InvariantCulture));
            }

            return builder.ToString();
        }

        private void ClearResults()
        {
            generatedNumbers.Clear();
            ResultsTextBox.Clear();
            ResultSummaryTextBlock.Text = "Aucun nombre généré";
            CopyButton.IsEnabled = false;
            ExportCsvButton.IsEnabled = false;
            HideMessages();
        }

        private void ShowError(string message)
        {
            StatusTextBlock.Visibility = Visibility.Collapsed;
            ValidationTextBlock.Text = message;
            ValidationTextBlock.Visibility = Visibility.Visible;
        }

        private void ShowStatus(string message)
        {
            ValidationTextBlock.Visibility = Visibility.Collapsed;
            StatusTextBlock.Text = message;
            StatusTextBlock.Visibility = Visibility.Visible;
        }

        private void HideMessages()
        {
            ValidationTextBlock.Visibility = Visibility.Collapsed;
            StatusTextBlock.Visibility = Visibility.Collapsed;
        }
    }
}
