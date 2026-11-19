using System.Windows;

namespace Docx2PDFLive.UI.Views;

public partial class FirstRunWelcomeWindow : Window
{
    public FirstRunWelcomeWindow()
    {
        InitializeComponent();
    }

    private void GetStartedButton_Click(object sender, RoutedEventArgs e)
    {
        DialogResult = true;
        Close();
    }
}
