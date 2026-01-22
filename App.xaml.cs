using System.Configuration;
using System.Data;
using System.Windows;

namespace MeineWPFApp;

public interface IApp
{
    static abstract void Main();
    void InitializeComponent();
}

/// <summary>
/// Interaction logic for App.xaml
/// </summary>
public partial class App : Application, IApp
{
}

