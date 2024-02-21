using IS.ViewModels;
using System.Windows;

namespace IS
{
    /// <summary>
    /// Interaction logic for App.xaml
    /// </summary>
    public partial class App : Application
    {
        protected override void OnStartup(StartupEventArgs e)
        {
            base.OnStartup(e);
            MainWindow window = new();

            var vm = new MainVewModel();

            window.DataContext = vm;

            window.Show();
        }
    }
}
