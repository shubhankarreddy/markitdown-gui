using System.Windows;
using System.Windows.Threading;

namespace MarkItDown.Native;

public partial class App : System.Windows.Application
{
    private void OnDispatcherUnhandledException(object sender, DispatcherUnhandledExceptionEventArgs e)
    {
        System.Windows.MessageBox.Show(
            e.Exception.Message,
            "MarkItDown",
            System.Windows.MessageBoxButton.OK,
            System.Windows.MessageBoxImage.Error);

        e.Handled = true;
    }
}
