using System;
using System.Windows;
using Serilog;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;

namespace seConfSW
{
    public partial class App : Application
    {
        private ServiceProvider _serviceProvider;
        public static IConfiguration Configuration { get; private set; }
        protected override void OnStartup(StartupEventArgs e)
        {
            base.OnStartup(e);

            // Настройка конфигурации
            Configuration = new ConfigurationBuilder()
                .SetBasePath(AppDomain.CurrentDomain.BaseDirectory)
                .AddJsonFile("appsettings.json", optional: false, reloadOnChange: true)
                .AddEnvironmentVariables()
                .Build();

            // Настройка Serilog
            Log.Logger = new LoggerConfiguration()
                .MinimumLevel.Information()
                .WriteTo.Console()
                .WriteTo.File("logs/log-.txt",
                    rollingInterval: RollingInterval.Day,
                    outputTemplate: "{Timestamp:yyyy-MM-dd HH:mm:ss} [{Level:u3}] {Message:lj}{NewLine}{Exception}")
                .CreateLogger();

            // Настройка DI
            var services = new ServiceCollection();
            //services.AddSingleton<IConfiguration>(configuration);
            services.AddSingleton<ILogger>(Log.Logger);

            


            _serviceProvider = services.BuildServiceProvider();

            try
            {
                Log.Information("Application started successfully.");
            }
            catch (Exception ex)
            {
                Log.Error(ex, "Error during application startup.");
                Shutdown(1);
            }
        }

        protected override void OnExit(ExitEventArgs e)
        {
            Log.Information("Application is shutting down.");
            Log.CloseAndFlush();
            base.OnExit(e);
        }
    }
}