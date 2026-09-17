using Microsoft.Extensions.DependencyInjection;
using sttmerve_cc.Services;
using sttmerve_cc.ViewModels;
using System;
using System.Windows;

namespace sttmerve_cc
{
    public partial class App : Application
    {
        public IServiceProvider ServiceProvider { get; private set; }

        protected override void OnStartup(StartupEventArgs e)
        {
            base.OnStartup(e);

            var services = new ServiceCollection();

            // 1. Singletons (Shared state)
            services.AddSingleton<CommandRepository>();
            services.AddSingleton<ConfigService>();

            // 2. Workers
            services.AddSingleton<ExcelService>(sp => {
                var config = sp.GetRequiredService<ConfigService>();
                return new ExcelService(config.ExcelFilePath, sp.GetRequiredService<CommandRepository>());
            });

            // Note: We pass the repository to the server so it can save immediately
            services.AddSingleton<HttpServerService>(sp => {
                var config = sp.GetRequiredService<ConfigService>();
                return new HttpServerService(
                    8080,
                    config,
                    sp.GetRequiredService<CommandRepository>(),
                    sp.GetRequiredService<ExcelService>()
                );
            });

            // 3. ViewModel
            services.AddTransient<MainViewModel>();

            ServiceProvider = services.BuildServiceProvider();
        }

    }
}
