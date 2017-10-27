using System.Windows;

namespace selenium_csharp
{
    /// <summary>
    /// Interaction logic for App.xaml
    /// </summary>
    public partial class App : Application
    {
        public App()
        {
            Dispatcher.UnhandledException += (s, e) =>
            {
                if (e.Exception.GetType() == typeof(OpenQA.Selenium.NoSuchElementException))
                {
                    e.Handled = true;
                    MessageBox.Show("Please check the network connection!", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                    Current.Shutdown();
                }
            };
        }
    }
}
