using System.Windows;
using Docx2PDFLive.Core.Services;

namespace Docx2PDFLive.UI.Views;

public partial class FirstRunWelcomeWindow : Window
{
    public FirstRunWelcomeWindow()
    {
        InitializeComponent();
    }

    private void GetStartedButton_Click(object sender, RoutedEventArgs e)
    {
        // Write the first-run marker to %LocalAppData%\Docx2PDFLive\welcomed.flag
        FirstRunMarker.MarkWelcomed();
        DialogResult = true;
        Close();
    }
}
