using System.Windows;
using System.Windows.Controls;
using MonTableurApp.Models;
using MonTableurApp.ViewModels;

namespace MonTableurApp.Views
{
    public partial class VueAjouterProjetView : UserControl
    {
        public VueAjouterProjetView()
        {
            InitializeComponent();
            Loaded += VueAjouterProjetView_Loaded;
        }

        private void VueAjouterProjetView_Loaded(object sender, RoutedEventArgs e)
        {
            if (DataContext is not MainViewModel viewModel)
            {
                return;
            }

            if (ClientComboBox.SelectedItem == null && viewModel.Clients.Count > 0)
            {
                ClientComboBox.SelectedItem = viewModel.Clients[0];
            }

            if (DemandeurComboBox.SelectedItem == null && viewModel.Demandeurs.Count > 0)
            {
                DemandeurComboBox.SelectedItem = viewModel.Demandeurs[0];
            }

            if (TypeActiviteComboBox.SelectedItem == null && viewModel.TypesActivite.Count > 0)
            {
                TypeActiviteComboBox.SelectedItem = viewModel.TypesActivite[0];
            }

            if (StatutComboBox.SelectedItem == null && viewModel.Statuts.Count > 0)
            {
                StatutComboBox.SelectedItem = viewModel.Statuts[0];
            }
        }

        private void AjouterProjet_Click(object sender, RoutedEventArgs e)
        {
            if (DataContext is not MainViewModel viewModel)
            {
                return;
            }

            string numeroProjet = NumeroProjetTextBox.Text.Trim();
            string nomProduit = NomProduitTextBox.Text.Trim();

            if (string.IsNullOrWhiteSpace(numeroProjet) || string.IsNullOrWhiteSpace(nomProduit))
            {
                MessageBox.Show(
                    "Le numéro projet et le nom produit sont obligatoires.",
                    "Projet incomplet",
                    MessageBoxButton.OK,
                    MessageBoxImage.Warning);
                return;
            }

            if (viewModel.ProjetExiste(numeroProjet, nomProduit))
            {
                MessageBox.Show(
                    "Un projet avec ce numéro et ce nom produit existe déjà.",
                    "Doublon détecté",
                    MessageBoxButton.OK,
                    MessageBoxImage.Information);
                return;
            }

            var projet = new Projet
            {
                NumeroProjet = numeroProjet,
                NomProduit = nomProduit,
                Client = ClientComboBox.SelectedItem as string,
                Demandeur = DemandeurComboBox.SelectedItem as string,
                TypeActivite = TypeActiviteComboBox.SelectedItem as string,
                DossierRacine = DossierRacineTextBox.Text.Trim(),
                Statut = StatutComboBox.SelectedItem as string,
                DateDebut = DateDebutTextBox.Text.Trim(),
                DatePrevisionnelle = DatePrevisionnelleTextBox.Text.Trim(),
                DateFin = DateFinTextBox.Text.Trim(),
                Commentaires = CommentairesTextBox.Text.Trim()
            };

            viewModel.AjouterProjet(projet);
            ViderFormulaire();

            MessageBox.Show(
                "Le projet a bien été ajouté.",
                "Projet créé",
                MessageBoxButton.OK,
                MessageBoxImage.Information);
        }

        private void ViderFormulaire_Click(object sender, RoutedEventArgs e)
        {
            ViderFormulaire();
        }

        private void ViderFormulaire()
        {
            NumeroProjetTextBox.Clear();
            NomProduitTextBox.Clear();
            DossierRacineTextBox.Clear();
            DateDebutTextBox.Clear();
            DatePrevisionnelleTextBox.Clear();
            DateFinTextBox.Clear();
            CommentairesTextBox.Clear();

            if (DataContext is not MainViewModel viewModel)
            {
                return;
            }

            ClientComboBox.SelectedItem = viewModel.Clients.Count > 0 ? viewModel.Clients[0] : null;
            DemandeurComboBox.SelectedItem = viewModel.Demandeurs.Count > 0 ? viewModel.Demandeurs[0] : null;
            TypeActiviteComboBox.SelectedItem = viewModel.TypesActivite.Count > 0 ? viewModel.TypesActivite[0] : null;
            StatutComboBox.SelectedItem = viewModel.Statuts.Count > 0 ? viewModel.Statuts[0] : null;
            NumeroProjetTextBox.Focus();
        }
    }
}
