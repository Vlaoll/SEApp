// Ignore Spelling: App Conf

using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using seConfSW.Presentation.ViewModels;
using seConfSW.Presentation.Views;
using seConfSW.Services;
using Serilog;
using System;
using System.IO;
using System.Windows;

namespace seConfSW
{
    public partial class App : Application
    {
        public static IServiceProvider ServiceProvider { get; private set; }


        public App()
        {
            var serviceCollection = new ServiceCollection();
            ConfigureServices(serviceCollection);
            ServiceProvider = serviceCollection.BuildServiceProvider();
        }

        private void ConfigureServices(IServiceCollection services)
        {
            // Настройка конфигурации
            var configuration = new ConfigurationBuilder()
                .SetBasePath(Directory.GetCurrentDirectory())
                .AddJsonFile("appsettings.json", optional: false, reloadOnChange: true)
                .Build();
            services.AddSingleton<IConfiguration>(configuration);

            // Настройка Serilog
            Log.Logger = new LoggerConfiguration()
                .MinimumLevel.Debug()
                .WriteTo.File($"log/Log_{DateTime.Now:yyyyMMdd_HHmmss}.txt",
                            rollingInterval: RollingInterval.Infinite, 
                            retainedFileCountLimit: null) 
                .ReadFrom.Configuration(configuration) 
                .CreateLogger();





            // Регистрация сервисов
            services.AddSingleton<ILogger>(Log.Logger);
            services.AddSingleton<IConfigurationService, Configuration>();
            services.AddSingleton<IExcelDataReader, ExcelDataReader>();
            services.AddSingleton<IExcelService, ExcelService>();
            services.AddSingleton<ITiaService, TiaService>();

            services.AddSingleton<IHierarchyManager, HierarchyManager>();
            services.AddSingleton<IPlcHardwareManager, PlcHardwareManager>();
            services.AddSingleton<IProjectManager, ProjectManager>();
            services.AddSingleton<ILibraryManager, LibraryManager>();
            services.AddSingleton<ICompilerManager, CompilerManager>();
            services.AddSingleton<ITagManager, TagManager>();
            services.AddSingleton<IPlcMasterCopyManager, PlcMasterCopyManager>();
            services.AddScoped<PlcBlockManager>();
            services.AddScoped<IPlcBlockManager>(sp => sp.GetRequiredService<PlcBlockManager>());
            services.AddScoped<IPlcSourceManager>(sp => sp.GetRequiredService<PlcBlockManager>());
            

            services.AddSingleton<MainWindowViewModel>();
            services.AddSingleton<MainWindow>();
        }

        protected override void OnStartup(StartupEventArgs e)
        {
            try
            {
                base.OnStartup(e);
                var mainWindow = ServiceProvider.GetRequiredService<MainWindow>();
                mainWindow.DataContext = ServiceProvider.GetRequiredService<MainWindowViewModel>();
                mainWindow.Show();
            }
            catch (Exception ex)
            {
                Log.Logger.Fatal(ex, "Application failed to start");
                MessageBox.Show($"Application failed to start: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                Shutdown(1);
            }
        }

        protected override void OnExit(ExitEventArgs e)
        {
            Log.CloseAndFlush(); // Закрытие Serilog перед выходом
            base.OnExit(e);
        }
    }
}