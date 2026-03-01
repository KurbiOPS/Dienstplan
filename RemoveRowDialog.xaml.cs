using System.Collections.ObjectModel;
using System.Windows;

namespace MeineWPFApp;

public partial class RemoveRowDialog : Window
{
    public int SelectedIndex { get; private set; } = -1;
    public bool Confirmed { get; private set; } = false;

    public RemoveRowDialog(ObservableCollection<string> names)
    {
        InitializeComponent();
        lstNames.ItemsSource = names;
    }

    private void btnRemove_Click(object sender, RoutedEventArgs e)
    {
        if (lstNames.SelectedIndex >= 0)
        {
            SelectedIndex = lstNames.SelectedIndex;
            Confirmed = true;
            this.DialogResult = true;
            this.Close();
        }
        else
        {
            MessageBox.Show("Bitte wählen Sie eine Person aus.", "Keine Auswahl", MessageBoxButton.OK, MessageBoxImage.Warning);
        }
    }

    private void btnCancel_Click(object sender, RoutedEventArgs e)
    {
        this.DialogResult = false;
        this.Close();
    }
}