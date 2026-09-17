using Microsoft.Extensions.DependencyInjection;
using sttmerve_cc.ViewModels;
using System.Windows;

namespace sttmerve_cc
{
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();

            // Ask the DI container to build the ViewModel and all its dependencies
            var app = (App)Application.Current;
            this.DataContext = app.ServiceProvider.GetRequiredService<MainViewModel>();
        }
    }
}
