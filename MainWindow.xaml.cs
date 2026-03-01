using System.Collections.ObjectModel;
using System.Diagnostics;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Xml.Serialization;
using static System.Net.Mime.MediaTypeNames;
using System.IO;
using System.Globalization;
using static System.Runtime.InteropServices.JavaScript.JSType;

namespace MeineWPFApp;

public partial class MainWindow : Window
{
    private readonly string[] Jahre = 
    { 
        "2026", "2027", "2028", "2029", "2030", "2031", 
        "2032", "2033", "2034", "2035", "2036", "2037" 
    };
    private readonly string[] monate = 
    { 
        "Januar", "Februar", "März", "April", "Mai", "Juni", 
        "Juli", "August", "September", "Oktober", "November", "Dezember" 
    };

    private int currentRowCount = 5; // Standardmäßig 5 Zeilen
    private List<TextBox> nameTextBoxes = new List<TextBox>();
    private List<TextBox> sollTextBoxes = new List<TextBox>();
    private List<TextBox> istTextBoxes = new List<TextBox>();
    private List<TextBox> percentTextBoxes = new List<TextBox>();
    private List<TextBox> dienstTextBoxes = new List<TextBox>();
    private List<TextBox> vmTextBoxes = new List<TextBox>();
    private List<TextBox> aktTextBoxes = new List<TextBox>();
    private ComboBox? cmbJahr;
    private ComboBox? cmbMonat;
    private TextBox? txtGlobalNote;
    
    private bool _isInitializing = true;

    public MainWindow()
    {

        InitializeComponent();
        LoadRowCount();
        GenerateCalendar();
        LoadAllFields();
        LoadGlobalNote();
        
        string selectedYearStr = cmbJahr?.SelectedItem?.ToString() ?? "2026";
        string selectedMonthStr = cmbMonat?.SelectedItem?.ToString() ?? "Januar";
        if (int.TryParse(selectedYearStr, out int year))
        {
            int monthIndex = Array.IndexOf(monate, selectedMonthStr);
            if (monthIndex >= 0) LoadDienstData(year, monthIndex + 1);
        }
        
        _isInitializing = false;
        this.Closing += MainWindow_Closing;
    }


    private void LoadRowCount()
    {
        try
        {
            string appDataPath = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            string settingsDir = System.IO.Path.Combine(appDataPath, "MeineWPFApp");
            string filePath = System.IO.Path.Combine(settingsDir, "settings_rowcount.txt");

            if (File.Exists(filePath))
            {
                string content = File.ReadAllText(filePath);
                if (int.TryParse(content, out int rowCount) && rowCount >= 1 && rowCount <= 20)
                {
                    currentRowCount = rowCount;
                }
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Fehler beim Laden der Zeilenanzahl: {ex.Message}");
            currentRowCount = 5;
        }
    }

    private void SaveRowCount()
    {
        try
        {
            string appDataPath = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            string settingsDir = System.IO.Path.Combine(appDataPath, "MeineWPFApp");
            Directory.CreateDirectory(settingsDir);
            string filePath = System.IO.Path.Combine(settingsDir, "settings_rowcount.txt");
            File.WriteAllText(filePath, currentRowCount.ToString());
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Fehler beim Speichern der Zeilenanzahl: {ex.Message}");
        }
    }

    private bool IsGreenDay(DateTime date)
    {
        bool isWeekend = date.DayOfWeek == DayOfWeek.Saturday || date.DayOfWeek == DayOfWeek.Sunday;
        return isWeekend || IsHoliday(date);
    }

    private bool IsHoliday(DateTime d)
    {
        if ((d.Month, d.Day) is (1, 1) or (5, 1) or (10, 3) or (12, 25) or (12, 26))
        return true;
        
        return false;
    }

    private readonly Brush GreenDayBackground = new SolidColorBrush(Color.FromRgb(220, 245, 220));
    private readonly Brush NormalDayBackground = Brushes.White;
    private readonly Brush TueThuNumberColor = Brushes.Blue;
    private readonly Brush HolidayNumberBackground =  new SolidColorBrush(Color.FromRgb(220, 245, 220));
    private readonly Brush NormalNumberBackground = Brushes.Black;

    private void cmbJahr_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (_isInitializing) return; // Verhindere Ausführung während Initialisierung

        LoadRowCount();
        GenerateCalendar();
        LoadAllFields();
        LoadGlobalNote();
        // Lade Dienst-Daten mit den neuen Werten
        string selectedYearStr = cmbJahr?.SelectedItem?.ToString() ?? "2026";
        string selectedMonthStr = cmbMonat?.SelectedItem?.ToString() ?? "Januar";
        if (int.TryParse(selectedYearStr, out int year))
        {
            int monthIndex = Array.IndexOf(monate, selectedMonthStr);
            if (monthIndex >= 0)
            {
                int month = monthIndex + 1;
                LoadDienstData(year, month);
            }
        }
    }

    private void cmbMonat_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (_isInitializing) return; // Verhindere Ausführung während Initialisierung

        LoadRowCount();
        GenerateCalendar();
        LoadAllFields();
        LoadGlobalNote();
        // Lade Dienst-Daten mit den neuen Werten
        string selectedYearStr = cmbJahr?.SelectedItem?.ToString() ?? "2026";
        string selectedMonthStr = cmbMonat?.SelectedItem?.ToString() ?? "Januar";
        if (int.TryParse(selectedYearStr, out int year))
        {
            int monthIndex = Array.IndexOf(monate, selectedMonthStr);
            if (monthIndex >= 0)
            {
                int month = monthIndex + 1;
                LoadDienstData(year, month);
            }
        }
    }
    private void txt_KeyDown(object sender, KeyEventArgs e)
    {
        // Speichern beim Enter entfernt
    }

    private void MainWindow_Closing(object? sender, System.ComponentModel.CancelEventArgs e)
    {
        var result = MessageBox.Show("Möchten Sie die Änderungen speichern?", "Speichern bestätigen", MessageBoxButton.YesNoCancel, MessageBoxImage.Question);
        if (result == MessageBoxResult.Yes)
        {
            SaveAllFields();
        }
        else if (result == MessageBoxResult.Cancel)
        {
            e.Cancel = true;
        }
        // Bei No: einfach schließen ohne speichern
    }

    private void btnSave_Click(object sender, RoutedEventArgs e)
    {
        if (!IsManualVmMonth())
        {
            LoadVMData(); // Automatisches VM aus Vormonat (außer Startmonat)
        }
        CalculateAkt(); // Berechne akt-Werte erneut, falls sich durch VM-Änderungen etwas geändert hat
        SaveAllFields(); // save Soll/Ist/akt Werte
    }

    private void btnAddRow_Click(object sender, RoutedEventArgs e)
    {
        if (currentRowCount < 20) // Maximum 20 Zeilen
        {
            currentRowCount++;
            SaveRowCount();
            GenerateCalendar();
        }
    }

    private void btnRemoveRow_Click(object sender, RoutedEventArgs e)
    {
        if (currentRowCount <= 1)
        {
            MessageBox.Show("Es muss mindestens eine Zeile vorhanden sein.", "Nicht möglich", MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }

        // Sammle alle Namen für das Dialogfenster
        ObservableCollection<string> names = new ObservableCollection<string>();
        for (int i = 0; i < currentRowCount; i++)
        {
            string name = i < nameTextBoxes.Count ? nameTextBoxes[i].Text : "";
            names.Add(string.IsNullOrWhiteSpace(name) ? $"Person {i + 1}" : name);
        }

        // Öffne Dialogfenster
        RemoveRowDialog dialog = new RemoveRowDialog(names);
        dialog.Owner = this;
        bool? result = dialog.ShowDialog();

        if (result == true && dialog.Confirmed)
        {
            int indexToRemove = dialog.SelectedIndex;

            // Entferne die Daten für diese Zeile
            RemoveRowData(indexToRemove);

            // Aktualisiere Zeilenanzahl
            currentRowCount--;
            SaveRowCount();

            // Regeneriere UI
            GenerateCalendar();
        }
    }

    private void SaveAllFields()
    {
        // Sammle Namen
        List<string> names = new List<string>();
        foreach (TextBox tb in nameTextBoxes)
        {
            names.Add(tb.Text);
        }
        // Fülle mit leeren Strings auf, falls weniger als 5 Namen
        while (names.Count < 5)
        {
            names.Add("");
        }

        // Speichere Namen global (unabhängig von Monat/Jahr)
        string namesData = string.Join("|", names);
        string appDataPath = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
        string settingsDir = System.IO.Path.Combine(appDataPath, "MeineWPFApp");
        Directory.CreateDirectory(settingsDir);
        string namesFile = System.IO.Path.Combine(settingsDir, "settings_names.txt");
        File.WriteAllText(namesFile, namesData);

        // Sammle Soll/Ist Werte
        List<string> sollValues = new List<string>();
        List<string> istValues = new List<string>();
        List<string> aktValues = new List<string>();
        foreach (TextBox tb in sollTextBoxes)
        {
            sollValues.Add(tb.Text);
        }
        foreach (TextBox tb in istTextBoxes)
        {
            istValues.Add(tb.Text);
        }

        // Berechne akt-Werte
        CalculateAkt();

        foreach (TextBox tb in aktTextBoxes)
        {
            aktValues.Add(tb.Text);
        }

        // Sammle Prozent-Werte
        List<string> percentValues = new List<string>();
        foreach (TextBox tb in percentTextBoxes)
        {
            percentValues.Add(tb.Text);
        }

        // Fülle mit leeren Strings auf, falls weniger Werte
        while (percentValues.Count < currentRowCount)
        {
            percentValues.Add("");
        }

        // Speichere Prozent-Werte global (unabhängig von Monat/Jahr)
        string percentData = string.Join("|", percentValues);
        string percentFile = System.IO.Path.Combine(settingsDir, "settings_percent.txt");
        File.WriteAllText(percentFile, percentData);

        string valuesData = string.Join("|", sollValues.Concat(istValues).Concat(aktValues));

        // Nutze das ausgewählte Jahr und Monat für den Dateinamen
        string selectedYear = cmbJahr.SelectedItem?.ToString() ?? "2026";
        string selectedMonth = cmbMonat.SelectedItem?.ToString() ?? "Januar";
        string valuesFile = System.IO.Path.Combine(settingsDir, $"settings_{selectedYear}_{selectedMonth}.txt");
        File.WriteAllText(valuesFile, valuesData);

        // VM wird nur im manuellen Startmonat separat gespeichert.
        if (IsManualVmMonth())
        {
            List<string> vmValues = new List<string>();
            foreach (TextBox tb in vmTextBoxes)
            {
                vmValues.Add(tb.Text);
            }
            string vmData = string.Join("|", vmValues);
            string vmFile = System.IO.Path.Combine(settingsDir, $"settings_vm_{selectedYear}_{selectedMonth}.txt");
            File.WriteAllText(vmFile, vmData);
        }

        // Speichere globale Notiz
        SaveGlobalNote();

        // Speichere Dienst-Daten
        SaveDienstData();
    }

    private void LoadAllFields()
    {
        try
        {
            string appDataPath = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            string settingsDir = System.IO.Path.Combine(appDataPath, "MeineWPFApp");

            // Lade Namen global
            string namesFile = System.IO.Path.Combine(settingsDir, "settings_names.txt");
            if (File.Exists(namesFile))
            {
                string namesData = File.ReadAllText(namesFile);
                string[] nameFields = namesData.Split('|');

                // Setze Namen in die dynamischen TextBoxen
                for (int i = 0; i < Math.Min(nameTextBoxes.Count, nameFields.Length); i++)
                {
                    nameTextBoxes[i].Text = nameFields[i];
                }
            }

            // Lade Prozent-Werte global
            string percentFile = System.IO.Path.Combine(settingsDir, "settings_percent.txt");
            if (File.Exists(percentFile))
            {
                string percentData = File.ReadAllText(percentFile);
                string[] percentFields = percentData.Split('|');

                // Setze Prozent-Werte in die dynamischen TextBoxen
                for (int i = 0; i < Math.Min(percentTextBoxes.Count, percentFields.Length); i++)
                {
                    percentTextBoxes[i].Text = percentFields[i];
                }
            }

            // Lade Soll/Ist Werte pro Monat und Jahr
            string selectedYear = cmbJahr.SelectedItem?.ToString() ?? "2026";
            string selectedMonth = cmbMonat.SelectedItem?.ToString() ?? "Januar";
            string valuesFile = System.IO.Path.Combine(settingsDir, $"settings_{selectedYear}_{selectedMonth}.txt");

            // Leere die Soll/Ist Felder zuerst
            foreach (TextBox tb in sollTextBoxes)
            {
                tb.Text = "";
            }
            foreach (TextBox tb in istTextBoxes)
            {
                tb.Text = "";
            }
            foreach (TextBox tb in aktTextBoxes)
            {
                tb.Text = "";
            }

            if (File.Exists(valuesFile))
            {
                string valuesData = File.ReadAllText(valuesFile);
                string[] valueFields = valuesData.Split('|');

                // Dateiformate:
                // Neu: Soll|Ist|akt (3 Blöcke)
                // Alt: Soll|Ist      (2 Blöcke)
                int storedRowCount = 0;
                bool hasAktBlock = false;

                if (valueFields.Length % 3 == 0 && valueFields.Length > 0)
                {
                    storedRowCount = valueFields.Length / 3;
                    hasAktBlock = true;
                }
                else if (valueFields.Length % 2 == 0 && valueFields.Length > 0)
                {
                    storedRowCount = valueFields.Length / 2;
                }
                else
                {
                    storedRowCount = currentRowCount;
                }

                int sollCount = storedRowCount;
                int istCount = storedRowCount;

                // Lade Soll-Werte
                for (int i = 0; i < Math.Min(sollTextBoxes.Count, storedRowCount); i++)
                {
                    if (i < valueFields.Length)
                    {
                        sollTextBoxes[i].Text = valueFields[i];
                    }
                }

                // Lade Ist-Werte
                for (int i = 0; i < Math.Min(istTextBoxes.Count, storedRowCount); i++)
                {
                    if (i + sollCount < valueFields.Length)
                    {
                        istTextBoxes[i].Text = valueFields[i + sollCount];
                    }
                }

                // Lade akt-Werte
                for (int i = 0; i < Math.Min(aktTextBoxes.Count, storedRowCount); i++)
                {
                    if (hasAktBlock && i + sollCount + istCount < valueFields.Length)
                    {
                        aktTextBoxes[i].Text = valueFields[i + sollCount + istCount];
                    }
                }
            }

            // Lade VM-Werte
            if (IsManualVmMonth())
            {
                LoadManualVmData();
            }
            else
            {
                LoadVMData(); // akt vom vorherigen Monat
            }

            // Berechne akt-Werte
            CalculateAkt();
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Fehler beim Laden: {ex.Message}");
        }
    }

    private void LoadGlobalNote()
    {
        try
        {
            string appDataPath = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            string settingsDir = System.IO.Path.Combine(appDataPath, "MeineWPFApp");
            string globalNoteFile = System.IO.Path.Combine(settingsDir, "settings_global_note.txt");
            if (File.Exists(globalNoteFile))
            {
                string globalNote = File.ReadAllText(globalNoteFile);
                txtGlobalNote.Text = globalNote;
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Fehler beim Laden der globalen Notiz: {ex.Message}");
        }
    }

    private void SaveGlobalNote()
    {
        try
        {
            string appDataPath = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            string settingsDir = System.IO.Path.Combine(appDataPath, "MeineWPFApp");
            Directory.CreateDirectory(settingsDir);
            string globalNoteFile = System.IO.Path.Combine(settingsDir, "settings_global_note.txt");
            File.WriteAllText(globalNoteFile, txtGlobalNote.Text);
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Fehler beim Speichern der globalen Notiz: {ex.Message}");
        }
    }

    private void LoadVMData()
    {
        try
        {
            string appDataPath = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            string settingsDir = System.IO.Path.Combine(appDataPath, "MeineWPFApp");

            string selectedYear = cmbJahr.SelectedItem?.ToString() ?? "2026";
            string selectedMonth = cmbMonat.SelectedItem?.ToString() ?? "Januar";

            // Berechne vorherigen Monat
            int year = int.Parse(selectedYear);
            int monthIndex = Array.IndexOf(monate, selectedMonth);
            int month = monthIndex + 1;
            month--;
            if (month == 0)
            {
                month = 12;
                year--;
            }
            string prevYear = year.ToString();
            string prevMonth = monate[month - 1];

            string prevValuesFile = System.IO.Path.Combine(settingsDir, $"settings_{prevYear}_{prevMonth}.txt");

            // Leere VM-Felder
            foreach (TextBox tb in vmTextBoxes)
            {
                tb.Text = "";
            }

            if (File.Exists(prevValuesFile))
            {
                string valuesData = File.ReadAllText(prevValuesFile);
                string[] valueFields = valuesData.Split('|');

                // Datei enthält Soll|Ist|akt mit gleicher Anzahl pro Block.
                // Nutze die tatsächliche Anzahl aus der Vor-Monats-Datei, damit geänderte Zeilenanzahl
                // (zwischen Monaten) die Offsets nicht verschiebt.
                int prevRowCount = valueFields.Length / 3;
                if (prevRowCount <= 0)
                {
                    return;
                }

                int sollCount = prevRowCount;
                int istCount = prevRowCount;
                int aktCount = prevRowCount;

                // Lade akt-Werte vom vorherigen Monat als VM
                for (int i = 0; i < Math.Min(vmTextBoxes.Count, aktCount); i++)
                {
                    if (i + sollCount + istCount < valueFields.Length)
                    {
                        vmTextBoxes[i].Text = valueFields[i + sollCount + istCount];
                    }
                }
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Fehler beim Laden der VM-Daten: {ex.Message}");
        }
    }

    private bool TryReadNumber(TextBox box, string fieldName, out double value)
    {
        value = ParseFlexibleDouble(box.Text);
        
        // Reject non-numeric text explicity
        string raw = box.Text?.Trim() ?? "";
        if (raw.Length > 0 &&
            !double.TryParse(raw, NumberStyles.Float, CultureInfo.CurrentCulture, out _)&& 
            !double.TryParse(raw, NumberStyles.Float, CultureInfo.InvariantCulture, out _))
        {
            MessageBox.Show($"{fieldName}: Bitte eine gültige Zahl eingeben.");
            box.Focus();
            return false;
        }

        return true;
    }

    private void CalculateAkt()
    {
        for (int i = 0; i < Math.Min(Math.Min(Math.Min(istTextBoxes.Count, sollTextBoxes.Count), vmTextBoxes.Count), aktTextBoxes.Count); i++)
        {

            if (!TryReadNumber(istTextBoxes[i], "Ist", out double ist)) return;
            if (!TryReadNumber(sollTextBoxes[i], "Soll", out double soll)) return;
            if (!TryReadNumber(vmTextBoxes[i], "VM", out double vm)) return;


            // Ist darf nur außerhalb des Startmonats nicht negativ sein.
            if (!IsManualVmMonth() && ist < 0)
            {
                ist = 0;
                istTextBoxes[i].Text = "0";
            }

            double akt = (ist + vm) - soll;
            aktTextBoxes[i].Text = akt.ToString("F2");
        }
    }

    private double ParseFlexibleDouble(string? input)
    {
        if (string.IsNullOrWhiteSpace(input))
        {
            return 0;
        }

        string value = input.Trim();

        // Erst CurrentCulture (de-DE: 7,5), dann Invariant (7.5).
        // Wichtig: Invariant zuerst würde "7,5" als 75 interpretieren.
        if (double.TryParse(value, NumberStyles.Float, CultureInfo.CurrentCulture, out double parsed))
        {
            return parsed;
        }

        if (double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out parsed))
        {
            return parsed;
        }

        // Fallback: Dezimaltrennzeichen vereinheitlichen.
        string normalized = value.Replace(',', '.');
        if (double.TryParse(normalized, NumberStyles.Float, CultureInfo.InvariantCulture, out parsed))
        {
            return parsed;
        }

        return 0;
    }

    private bool IsManualVmMonth()
    {
        string selectedYear = cmbJahr?.SelectedItem?.ToString() ?? "2026";
        string selectedMonth = cmbMonat?.SelectedItem?.ToString() ?? "Januar";
        return selectedYear == "2026" && selectedMonth == "Januar";
    }

    private void LoadManualVmData()
    {
        try
        {
            string selectedYear = cmbJahr?.SelectedItem?.ToString() ?? "2026";
            string selectedMonth = cmbMonat?.SelectedItem?.ToString() ?? "Januar";
            string appDataPath = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            string settingsDir = System.IO.Path.Combine(appDataPath, "MeineWPFApp");
            string vmFile = System.IO.Path.Combine(settingsDir, $"settings_vm_{selectedYear}_{selectedMonth}.txt");

            foreach (TextBox tb in vmTextBoxes)
            {
                tb.Text = "";
            }

            if (File.Exists(vmFile))
            {
                string vmData = File.ReadAllText(vmFile);
                string[] vmFields = vmData.Split('|');
                for (int i = 0; i < Math.Min(vmTextBoxes.Count, vmFields.Length); i++)
                {
                    vmTextBoxes[i].Text = vmFields[i];
                }
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Fehler beim Laden der manuellen VM-Daten: {ex.Message}");
        }
    }

    private void GenerateCalendar()
    {
        try
        {
            // Leere die dynamischen Elemente
            nameTextBoxes.Clear();
            sollTextBoxes.Clear();
            istTextBoxes.Clear();
            percentTextBoxes.Clear();
            dienstTextBoxes.Clear();
            vmTextBoxes.Clear();
            aktTextBoxes.Clear();
            mainGrid.RowDefinitions.Clear();
            mainGrid.ColumnDefinitions.Clear();
            mainGrid.Children.Clear();

            string selectedYearStr = cmbJahr?.SelectedItem?.ToString() ?? "2026";
            string selectedMonthStr = cmbMonat?.SelectedItem?.ToString() ?? "Januar";

            if (int.TryParse(selectedYearStr, out int year))
            {
                int monthIndex = Array.IndexOf(monate, selectedMonthStr);
                if (monthIndex >= 0)
                {
                    int month = monthIndex + 1; // Monate sind 1-basiert
                    int daysInMonth = DateTime.DaysInMonth(year, month);

                    // Lade Namen zuerst, um die maximale Breite zu berechnen
                    List<string> currentNames = LoadNamesForWidthCalculation();
                    double nameColumnWidth = CalculateNameColumnWidth(currentNames);

                    // Erstelle RowDefinitions: Jahr + Header + (Name + Prozent) pro Person
                    mainGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto }); // Jahr row
                    mainGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto }); // Header row
                    for (int i = 0; i < currentRowCount * 2; i++) // Name + Prozent pro Person
                    {
                        mainGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
                    }

                    // Erstelle ColumnDefinitions: Namen, Soll, Ist + Tage
                    mainGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(nameColumnWidth) }); // Namen (0)
                    mainGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(55) }); // Soll (1)
                    mainGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(55) });  // Ist (2)

                    for (int day = 1; day <= daysInMonth; day++)
                    {
                        mainGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(23) });
                    }

                    mainGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(55) }); // VM
                    mainGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(55) }); // akt

                    // Erstelle statische Header
                    AddStaticHeaders(nameColumnWidth, daysInMonth, selectedYearStr, selectedMonthStr);

                    bool isManualVmMonth = selectedYearStr == "2026" && selectedMonthStr == "Januar";

                    // Erstelle dynamische Felder für alle Personen (abwechselnd Name-Zeile und Prozent-Zeile)
                    for (int personIndex = 0; personIndex < currentRowCount; personIndex++)
                    {
                        int nameRowIndex = personIndex * 2 + 2; // Name-Zeile für diese Person (startet bei Row 2)
                        int percentRowIndex = personIndex * 2 + 3; // Prozent-Zeile für diese Person

                        // Name Field
                        Border nameBorder = new Border
                        {
                            BorderBrush = Brushes.LightGray,
                            BorderThickness = new Thickness(1),
                            Margin = new Thickness(0)
                        };
                        TextBox nameBox = new TextBox
                        {
                            Width = nameColumnWidth - 12, // -12 für Border
                            Height = 24,
                            Margin = new Thickness(0),
                            Padding = new Thickness(2),
                            FontSize = 14,
                            BorderThickness = new Thickness(0),
                            Background = Brushes.Transparent,
                            HorizontalContentAlignment = HorizontalAlignment.Center
                        };
                        nameBox.KeyDown += txt_KeyDown;
                        nameBorder.Child = nameBox;
                        Grid.SetRow(nameBorder, nameRowIndex);
                        Grid.SetColumn(nameBorder, 0);
                        mainGrid.Children.Add(nameBorder);
                        nameTextBoxes.Add(nameBox);

                        // Soll Field
                        Border sollBorder = new Border
                        {
                            BorderBrush = Brushes.LightGray,
                            BorderThickness = new Thickness(1),
                            Background = Brushes.White,
                            Margin = new Thickness(0)
                        };
                        TextBox sollBox = new TextBox
                        {
                            Width = 53,
                            Height = 24,
                            Margin = new Thickness(0),
                            Padding = new Thickness(2),
                            FontSize = 14,
                            BorderThickness = new Thickness(0),
                            Background = Brushes.Transparent,
                            HorizontalContentAlignment = HorizontalAlignment.Center
                        };
                        sollBox.KeyDown += txt_KeyDown;
                        sollBorder.Child = sollBox;
                        Grid.SetRow(sollBorder, nameRowIndex);
                        Grid.SetColumn(sollBorder, 1);
                        mainGrid.Children.Add(sollBorder);
                        sollTextBoxes.Add(sollBox);

                        // Ist Field
                        Border istBorder = new Border
                        {
                            BorderBrush = Brushes.LightGray,
                            BorderThickness = new Thickness(1),
                            Background = Brushes.White,
                            Margin = new Thickness(0)
                        };
                        TextBox istBox = new TextBox
                        {
                            Width = 53,
                            Height = 24,
                            Margin = new Thickness(0),
                            Padding = new Thickness(2),
                            FontSize = 14,
                            BorderThickness = new Thickness(0),
                            Background = Brushes.Transparent,
                            HorizontalContentAlignment = HorizontalAlignment.Center
                        };
                        istBox.KeyDown += txt_KeyDown;
                        istBorder.Child = istBox;
                        Grid.SetRow(istBorder, nameRowIndex);
                        Grid.SetColumn(istBorder, 2);
                        mainGrid.Children.Add(istBorder);
                        istTextBoxes.Add(istBox);

                        // VM Field
                        Border vmBorder = new Border
                        {
                            BorderBrush = Brushes.LightGray,
                            BorderThickness = new Thickness(1),
                            Background = Brushes.White,
                            Margin = new Thickness(0)
                        };
                        TextBox vmBox = new TextBox
                        {
                            Width = 53,
                            Height = 24,
                            Margin = new Thickness(0),
                            Padding = new Thickness(2),
                            FontSize = 14,
                            BorderThickness = new Thickness(0),
                            Background = isManualVmMonth ? Brushes.Transparent : Brushes.LightGray,
                            IsReadOnly = !isManualVmMonth,
                            HorizontalContentAlignment = HorizontalAlignment.Center
                        };
                        vmBox.KeyDown += txt_KeyDown;
                        vmBorder.Child = vmBox;
                        Grid.SetRow(vmBorder, nameRowIndex);
                        Grid.SetColumn(vmBorder, 3 + daysInMonth);
                        mainGrid.Children.Add(vmBorder);
                        vmTextBoxes.Add(vmBox);

                        // akt Field
                        Border aktBorder = new Border
                        {
                            BorderBrush = Brushes.LightGray,
                            BorderThickness = new Thickness(1),
                            Background = Brushes.White,
                            Margin = new Thickness(0)
                        };
                        TextBox aktBox = new TextBox
                        {
                            Width = 53,
                            Height = 24,
                            Margin = new Thickness(0),
                            Padding = new Thickness(2),
                            FontSize = 14,
                            BorderThickness = new Thickness(0),
                            Background = Brushes.Transparent,
                            HorizontalContentAlignment = HorizontalAlignment.Center
                        };
                        aktBox.KeyDown += txt_KeyDown;
                        aktBorder.Child = aktBox;
                        Grid.SetRow(aktBorder, nameRowIndex);
                        Grid.SetColumn(aktBorder, 4 + daysInMonth);
                        mainGrid.Children.Add(aktBorder);
                        aktTextBoxes.Add(aktBox);

                        // Prozent Field (nur in Namensspalte)
                        Border percentBorder = new Border
                        {
                            BorderBrush = Brushes.LightGray,
                            BorderThickness = new Thickness(1),
                            Background = Brushes.LightGray,
                            Margin = new Thickness(0)
                        };
                        TextBox percentBox = new TextBox
                        {
                            Width = nameColumnWidth - 12, // Gleiche Breite wie Namen
                            Height = 24,
                            Margin = new Thickness(0),
                            Padding = new Thickness(2),
                            FontSize = 14,
                            BorderThickness = new Thickness(0),
                            Background = Brushes.Transparent,
                            HorizontalContentAlignment = HorizontalAlignment.Center
                        };
                        percentBox.KeyDown += txt_KeyDown;
                        percentBorder.Child = percentBox;
                        // Event Handler entfernt
                        Grid.SetRow(percentBorder, percentRowIndex);
                        Grid.SetColumn(percentBorder, 0);
                        mainGrid.Children.Add(percentBorder);
                        percentTextBoxes.Add(percentBox);
                        // Leere Zellen in Prozent-Zeile (Soll, Ist, Tage, VM, akt) mit grauer Hintergrundfarbe
                        for (int col = 1; col <= 4 + daysInMonth; col++)
                        {
                            Border emptyBorder = new Border
                            {
                                BorderBrush = Brushes.LightGray,
                                BorderThickness = new Thickness(1),
                                Margin = new Thickness(0)
                            };

                            if (col >= 3 && col <= 2 + daysInMonth)
                            {
                                int dayNumber = col - 2;
                                DateTime d = new DateTime(year, month, dayNumber);
                                emptyBorder.Background = IsGreenDay(d) ? GreenDayBackground : Brushes.LightGray;
                            }
                            else
                            {
                                emptyBorder.Background = Brushes.LightGray;
                            }

                            Grid.SetRow(emptyBorder, percentRowIndex);
                            Grid.SetColumn(emptyBorder, col);
                            mainGrid.Children.Add(emptyBorder);
                        }
                    }
                    for (int day = 1; day <= daysInMonth; day++)
                    {
                        int columnIndex = 2 + day; // Start bei Spalte 3

                        // Berechne den Wochentag für diesen Tag
                        DateTime currentDate = new DateTime(year, month, day);
                        string shortDayName = GetShortDayName(currentDate.DayOfWeek);

                        bool isGreenDay = IsGreenDay(currentDate);
                        bool isTueThu = currentDate.DayOfWeek == DayOfWeek.Tuesday || currentDate.DayOfWeek == DayOfWeek.Thursday;

                        Brush tueThuHeader = new SolidColorBrush(Color.FromRgb(220, 235, 255)); // Heller Blauton für Di/Do
                        Brush greenHeader = GreenDayBackground; // Grün für grüne Tage
                        Brush normalHeader = Brushes.White; // Weiß für normale Tage

                        if (isGreenDay)
                        {
                            // Grün für grüne Tage
                            normalHeader = greenHeader;
                        }
                        else if (isTueThu)
                        {
                            // Heller Blauton für Di/Do
                            normalHeader = tueThuHeader;
                        }


                        Border headerBorder = new Border
                        {
                            BorderBrush = Brushes.Gray,
                            BorderThickness = new Thickness(1),
                            Margin = new Thickness(0)
                        };

                        if (isGreenDay)
                        {
                            headerBorder.Background = GreenDayBackground;
                        }
                        else if (isTueThu)
                        {
                            headerBorder.Background = Brushes.LightBlue;
                        }
                        else
                        {
                            headerBorder.Background = Brushes.White;
                        }

                        StackPanel headerPanel = new StackPanel
                        {
                            Orientation = Orientation.Vertical,
                            HorizontalAlignment = HorizontalAlignment.Center,
                            VerticalAlignment = VerticalAlignment.Center,
                            Margin = new Thickness(2)
                        };

                        TextBlock dayOfWeekText = new TextBlock
                        {
                            Text = shortDayName,
                            FontSize = 13,
                            FontWeight = FontWeights.Bold,
                            HorizontalAlignment = HorizontalAlignment.Center,
                            Foreground = Brushes.Gray
                        };

                        TextBlock dayNumberText = new TextBlock
                        {
                            Text = day.ToString(),
                            FontSize = 15,
                            FontWeight = FontWeights.Bold,
                            HorizontalAlignment = HorizontalAlignment.Center
                        };

                        dayNumberText.Foreground = isGreenDay ? Brushes.DarkGreen : Brushes.Black;
                        
                        headerPanel.Children.Add(dayOfWeekText);
                        headerPanel.Children.Add(dayNumberText);
                        headerBorder.Child = headerPanel;

                        Grid.SetRow(headerBorder, 1);
                        Grid.SetColumn(headerBorder, columnIndex);
                        mainGrid.Children.Add(headerBorder);

                        // Dienst-Felder für jede Person (nur in Name-Zeilen)
                        for (int personIndex = 0; personIndex < currentRowCount; personIndex++)
                        {
                            int nameRowIndex = personIndex * 2 + 2; // Name-Zeile für diese Person

                            Border dienstBorder = new Border
                            {
                                BorderBrush = Brushes.LightGray,
                                BorderThickness = new Thickness(1),
                                Background = Brushes.White,
                                Margin = new Thickness(0)
                            };
                            TextBox dienstBox = new TextBox
                            {
                                Width = 23,
                                Height = 24,
                                Margin = new Thickness(0),
                                Padding = new Thickness(1),
                                FontSize = 11,
                                BorderThickness = new Thickness(0),
                                Background = Brushes.Transparent,
                                HorizontalContentAlignment = HorizontalAlignment.Center,
                                VerticalContentAlignment = VerticalAlignment.Center
                            };
                            dienstBox.KeyDown += txt_KeyDown;
                            dienstBox.Tag = $"{personIndex + 1}_{day}"; // Tag für Identifikation
                            dienstBorder.Child = dienstBox;
                            dienstBorder.Background = isGreenDay ? GreenDayBackground : NormalDayBackground;
                            dienstBox.Background = Brushes.Transparent;

                            Grid.SetRow(dienstBorder, nameRowIndex);
                            Grid.SetColumn(dienstBorder, columnIndex);
                            mainGrid.Children.Add(dienstBorder);
                            dienstTextBoxes.Add(dienstBox);
                        }
                    }

                    // VM Header
                    Border vmHeaderBorder = new Border
                    {
                        BorderBrush = Brushes.Gray,
                        BorderThickness = new Thickness(1),
                        Margin = new Thickness(0)
                    };
                    TextBlock vmHeader = new TextBlock
                    {
                        Text = "VM",
                        FontSize = 16,
                        FontWeight = FontWeights.Bold,
                        Margin = new Thickness(5),
                        HorizontalAlignment = HorizontalAlignment.Center,
                        VerticalAlignment = VerticalAlignment.Center
                    };
                    vmHeaderBorder.Child = vmHeader;
                    Grid.SetRow(vmHeaderBorder, 1);
                    Grid.SetColumn(vmHeaderBorder, 3 + daysInMonth);
                    mainGrid.Children.Add(vmHeaderBorder);

                    // akt Header
                    Border aktHeaderBorder = new Border
                    {
                        BorderBrush = Brushes.Gray,
                        BorderThickness = new Thickness(1),
                        Margin = new Thickness(0)
                    };
                    TextBlock aktHeader = new TextBlock
                    {
                        Text = "akt",
                        FontSize = 16,
                        FontWeight = FontWeights.Bold,
                        Margin = new Thickness(5),
                        HorizontalAlignment = HorizontalAlignment.Center,
                        VerticalAlignment = VerticalAlignment.Center
                    };
                    aktHeaderBorder.Child = aktHeader;
                    Grid.SetRow(aktHeaderBorder, 1);
                    Grid.SetColumn(aktHeaderBorder, 4 + daysInMonth);
                    mainGrid.Children.Add(aktHeaderBorder);
                }
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Fehler beim Generieren des Dienstplans: {ex.Message}\n\nStackTrace: {ex.StackTrace}");
        }
    }

    private List<string> LoadNamesForWidthCalculation()
    {
        List<string> names = new List<string>();
        try
        {
            string appDataPath = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            string settingsDir = System.IO.Path.Combine(appDataPath, "MeineWPFApp");

            // Lade Namen global
            string namesFile = System.IO.Path.Combine(settingsDir, "settings_names.txt");
            if (File.Exists(namesFile))
            {
                string namesData = File.ReadAllText(namesFile);
                string[] nameFields = namesData.Split('|');

                // Füge alle Namen hinzu, auch leere
                for (int i = 0; i < Math.Max(currentRowCount, nameFields.Length); i++)
                {
                    if (i < nameFields.Length)
                    {
                        names.Add(nameFields[i]);
                    }
                    else
                    {
                        names.Add("");
                    }
                }
            }
            else
            {
                // Wenn keine Datei existiert, erstelle leere Namen
                for (int i = 0; i < currentRowCount; i++)
                {
                    names.Add("");
                }
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Fehler beim Laden der Namen für Breitenberechnung: {ex.Message}");
            // Fallback: leere Namen
            for (int i = 0; i < currentRowCount; i++)
            {
                names.Add("");
            }
        }
        return names;
    }

    private double CalculateNameColumnWidth(List<string> names)
    {
        // Mindestbreite
        double minWidth = 160;
        double maxWidth = 350; // Maximale Breite, um nicht zu breit zu werden

        if (names.Count == 0)
            return minWidth;

        // Finde den längsten Namen
        int maxLength = 0;
        foreach (string name in names)
        {
            if (name.Length > maxLength)
            {
                maxLength = name.Length;
            }
        }

        // Berechne Breite basierend auf Zeichenlänge (ca. 9 Pixel pro Zeichen bei FontSize 14)
        double calculatedWidth = Math.Max(minWidth, maxLength * 9 + 50); // +50 für Padding und Rand

        // Begrenze auf maximale Breite
        return Math.Min(calculatedWidth, maxWidth);
    }

    private void AddStaticHeaders(double nameColumnWidth, int daysInMonth, string selectedYearStr, string selectedMonthStr)
    {
        // Jahr ComboBox in Row 0, Column 0
        Border jahrBorder = new Border
        {
            BorderBrush = Brushes.Gray,
            BorderThickness = new Thickness(1),

            Margin = new Thickness(0)
        };
        cmbJahr = new ComboBox
        {
            Width = nameColumnWidth - 12,
            Height = 28,
            SelectedIndex = 0,
            FontSize = 14,
            Margin = new Thickness(2)
        };
        cmbJahr.ItemsSource = Jahre;
        cmbJahr.SelectedItem = selectedYearStr;
        cmbJahr.SelectionChanged += cmbJahr_SelectionChanged;
        jahrBorder.Child = cmbJahr;
        Grid.SetRow(jahrBorder, 0);
        Grid.SetColumn(jahrBorder, 0);
        mainGrid.Children.Add(jahrBorder);

        // Globales Notiz-Feld in Row 0, Column 1 mit ColSpan über die restlichen Spalten
        Border noteBorder = new Border
        {
            BorderBrush = Brushes.Gray,
            BorderThickness = new Thickness(1),
            Background = Brushes.LightYellow,
            Margin = new Thickness(0)
        };
        txtGlobalNote = new TextBox
        {
            Height = 28,
            Margin = new Thickness(2),
            FontSize = 18,
            FontWeight = FontWeights.Bold,
            HorizontalContentAlignment = HorizontalAlignment.Center,
            Background = Brushes.LightYellow,
            BorderThickness = new Thickness(0)
        };
        txtGlobalNote.KeyDown += txt_KeyDown;
        noteBorder.Child = txtGlobalNote;
        Grid.SetRow(noteBorder, 0);
        Grid.SetColumn(noteBorder, 1);
        Grid.SetColumnSpan(noteBorder, 4 + daysInMonth); // Von Column 1 bis Ende
        mainGrid.Children.Add(noteBorder);

        // Leere Zellen in Jahr-Zeile werden vom Notiz-Feld überdeckt

        // Monat ComboBox in Row 1, Column 0 (ersetzt Namen Header)
        Border monatBorder = new Border
        {
            BorderBrush = Brushes.Gray,
            BorderThickness = new Thickness(1),
            Margin = new Thickness(0)
        };
        cmbMonat = new ComboBox
        {
            Width = nameColumnWidth - 12,
            Height = 28,
            SelectedIndex = 0,
            FontSize = 14,
            Margin = new Thickness(2)
        };
        cmbMonat.ItemsSource = monate;
        cmbMonat.SelectedItem = selectedMonthStr;
        cmbMonat.SelectionChanged += cmbMonat_SelectionChanged;
        monatBorder.Child = cmbMonat;
        Grid.SetRow(monatBorder, 1);
        Grid.SetColumn(monatBorder, 0);
        mainGrid.Children.Add(monatBorder);

        // Soll Header in Row 1, Column 1
        Border sollHeaderBorder = new Border
        {
            BorderBrush = Brushes.Gray,
            BorderThickness = new Thickness(1),
            Margin = new Thickness(0)
        };
        TextBlock sollHeader = new TextBlock
        {
            Text = "Soll",
            FontSize = 16,
            FontWeight = FontWeights.Bold,
            Margin = new Thickness(5),
            HorizontalAlignment = HorizontalAlignment.Center,
            VerticalAlignment = VerticalAlignment.Center
        };
        sollHeaderBorder.Child = sollHeader;
        Grid.SetRow(sollHeaderBorder, 1);
        Grid.SetColumn(sollHeaderBorder, 1);
        mainGrid.Children.Add(sollHeaderBorder);

        // Ist Header in Row 1, Column 2
        Border istHeaderBorder = new Border
        {
            BorderBrush = Brushes.Gray,
            BorderThickness = new Thickness(1),
            Margin = new Thickness(0)
        };
        TextBlock istHeader = new TextBlock
        {
            Text = "Ist",
            FontSize = 16,
            FontWeight = FontWeights.Bold,
            Margin = new Thickness(5),
            HorizontalAlignment = HorizontalAlignment.Center,
            VerticalAlignment = VerticalAlignment.Center
        };
        istHeaderBorder.Child = istHeader;
        Grid.SetRow(istHeaderBorder, 1);
        Grid.SetColumn(istHeaderBorder, 2);
        mainGrid.Children.Add(istHeaderBorder);
    }



    private void SaveDienstData()
    {
        try
        {
            string selectedYear = cmbJahr.SelectedItem?.ToString() ?? "2026";
            string selectedMonth = cmbMonat.SelectedItem?.ToString() ?? "Januar";
            
            string appDataPath = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            string settingsDir = System.IO.Path.Combine(appDataPath, "MeineWPFApp");
            Directory.CreateDirectory(settingsDir);
            
            // Sammle alle Dienst-Daten
            List<string> dienstData = new List<string>();
            foreach (TextBox tb in dienstTextBoxes)
            {
                dienstData.Add(tb.Text);
            }
            
            // Speichere als Pipe-separierte Liste
            string data = string.Join("|", dienstData);
            string fileName = $"dienst_{selectedYear}_{selectedMonth}.txt";
            string filePath = System.IO.Path.Combine(settingsDir, fileName);
            File.WriteAllText(filePath, data);
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Fehler beim Speichern der Dienst-Daten: {ex.Message}");
        }
    }

    private string GetShortDayName(DayOfWeek dayOfWeek)
    {
        return dayOfWeek switch
        {
            DayOfWeek.Monday => "Mo",
            DayOfWeek.Tuesday => "Di",
            DayOfWeek.Wednesday => "Mi",
            DayOfWeek.Thursday => "Do",
            DayOfWeek.Friday => "Fr",
            DayOfWeek.Saturday => "Sa",
            DayOfWeek.Sunday => "So",
            _ => ""
        };
    }

    private void RemoveRowData(int indexToRemove)
    {
        try
        {
            string selectedYear = cmbJahr.SelectedItem?.ToString() ?? "2026";
            string selectedMonth = cmbMonat.SelectedItem?.ToString() ?? "Januar";

            // 1. Namen aktualisieren
            List<string> names = new List<string>();
            for (int i = 0; i < currentRowCount; i++)
            {
                if (i != indexToRemove)
                {
                    string name = i < nameTextBoxes.Count ? nameTextBoxes[i].Text : "";
                    names.Add(name);
                }
            }
            // Fülle mit leeren Strings auf, falls weniger als 5 Namen
            while (names.Count < 5)
            {
                names.Add("");
            }
            string namesData = string.Join("|", names);
            string appDataPath = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            string settingsDir = System.IO.Path.Combine(appDataPath, "MeineWPFApp");
            Directory.CreateDirectory(settingsDir);
            string namesFile = System.IO.Path.Combine(settingsDir, "settings_names.txt");
            File.WriteAllText(namesFile, namesData);

            // 1.5. Prozent-Werte aktualisieren
            List<string> percentValues = new List<string>();
            for (int i = 0; i < currentRowCount; i++)
            {
                if (i != indexToRemove)
                {
                    string percent = i < percentTextBoxes.Count ? percentTextBoxes[i].Text : "";
                    percentValues.Add(percent);
                }
            }
            string percentData = string.Join("|", percentValues);
            string percentFile = System.IO.Path.Combine(settingsDir, "settings_percent.txt");
            File.WriteAllText(percentFile, percentData);

            // 2. Soll/Ist Werte aktualisieren
            List<string> sollValues = new List<string>();
            List<string> istValues = new List<string>();
            List<string> aktValues = new List<string>();
            for (int i = 0; i < currentRowCount; i++)
            {
                if (i != indexToRemove)
                {
                    string soll = i < sollTextBoxes.Count ? sollTextBoxes[i].Text : "0";
                    string ist = i < istTextBoxes.Count ? istTextBoxes[i].Text : "0";
                    string akt = i < aktTextBoxes.Count ? aktTextBoxes[i].Text : "0";
                    sollValues.Add(soll);
                    istValues.Add(ist);
                    aktValues.Add(akt);
                }
            }
            string valuesData = string.Join("|", sollValues.Concat(istValues).Concat(aktValues));
            string valuesFile = System.IO.Path.Combine(settingsDir, $"settings_{selectedYear}_{selectedMonth}.txt");
            File.WriteAllText(valuesFile, valuesData);

            // 3. Dienst-Daten aktualisieren
            if (File.Exists(System.IO.Path.Combine(settingsDir, $"dienst_{selectedYear}_{selectedMonth}.txt")))
            {
                string dienstFile = System.IO.Path.Combine(settingsDir, $"dienst_{selectedYear}_{selectedMonth}.txt");
                string dienstData = File.ReadAllText(dienstFile);
                string[] dienstFields = dienstData.Split('|');

                // Berechne, wie viele Tage im Monat sind
                int year = int.Parse(selectedYear);
                int monthIndex = Array.IndexOf(monate, selectedMonth);
                int month = monthIndex + 1;
                int daysInMonth = DateTime.DaysInMonth(year, month);

                List<string> updatedDienstData = new List<string>();
                for (int day = 1; day <= daysInMonth; day++)
                {
                    for (int row = 0; row < currentRowCount; row++)
                    {
                        if (row != indexToRemove)
                        {
                            int index = (day - 1) * currentRowCount + row;
                            if (index < dienstFields.Length)
                            {
                                updatedDienstData.Add(dienstFields[index]);
                            }
                            else
                            {
                                updatedDienstData.Add("");
                            }
                        }
                    }
                }

                string updatedData = string.Join("|", updatedDienstData);
                File.WriteAllText(dienstFile, updatedData);
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Fehler beim Entfernen der Zeile: {ex.Message}");
        }
    }

    private void LoadDienstData(int year, int month)
    {
        try
        {
            if (month < 1 || month > 12)
            {
                MessageBox.Show($"Ungültiger Monat: {month}");
                return;
            }
            string appDataPath = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            string settingsDir = System.IO.Path.Combine(appDataPath, "MeineWPFApp");
            string monthName = monate[month - 1]; // Monat-Index zu Name konvertieren
            string fileName = $"dienst_{year}_{monthName}.txt";
            string filePath = System.IO.Path.Combine(settingsDir, fileName);
            
            if (File.Exists(filePath))
            {
                string data = File.ReadAllText(filePath);
                string[] dienstEntries = data.Split('|');
                
                // Setze die Daten in die TextBoxes
                for (int i = 0; i < Math.Min(dienstEntries.Length, dienstTextBoxes.Count); i++)
                {
                    dienstTextBoxes[i].Text = dienstEntries[i];
                }
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Fehler beim Laden der Dienst-Daten: {ex.Message}");
        }
    }
}
