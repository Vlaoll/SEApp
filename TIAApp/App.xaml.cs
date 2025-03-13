using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using seConfSW.Presentation.ViewModels;
using seConfSW.Presentation.Views;
using seConfSW.Services;
using System;
using System.IO;
using System.Windows;

namespace seConfSW
{
    public partial class App : Application
    {
        private readonly IServiceProvider _serviceProvider;
        public static IConfiguration Configuration { get; private set; }

        public App()
        {
            var serviceCollection = new ServiceCollection();
            ConfigureServices(serviceCollection);
            _serviceProvider = serviceCollection.BuildServiceProvider();
        }

        private void ConfigureServices(IServiceCollection services)
        {
            Configuration = new ConfigurationBuilder()
                .SetBasePath(Directory.GetCurrentDirectory())
                .AddJsonFile("appsettings.json", optional: false, reloadOnChange: true)
                .Build();
            services.AddSingleton<IConfiguration>(Configuration);

            services.AddSingleton<ExcelService>();
            services.AddSingleton<TiaService>();
            services.AddTransient<MainWindowViewModel>();
            services.AddTransient<MainWindow>();
        }

        protected override void OnStartup(StartupEventArgs e)
        {
            base.OnStartup(e);
            var mainWindow = _serviceProvider.GetRequiredService<MainWindow>();
            mainWindow.Show();
        }
    }
}