// Ignore Spelling: Conf plc

using Microsoft.Win32;
using seConfSW.Services;
using System;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Runtime.CompilerServices;
using System.Windows;
using System.Windows.Input;
using Serilog;
using System.Diagnostics;
using System.Linq;

namespace seConfSW.Presentation.ViewModels
{
    /// <summary>
    /// Represents the possible states of the project execution workflow
    /// </summary>
    public enum ProjectState
    {
        Initial,
        PermissionsChecked,
        ExcelRead,
        TiaConnected,
        LibrarySelected,
        Executing,
        Completed
    }

    /// <summary>
    /// Main ViewModel class for the application, handling the business logic and UI interactions
    /// </summary>
    public class MainWindowViewModel : INotifyPropertyChanged
    {
        #region Constants
        private const string LogPrefix = "[Main]";
        private const string LogFolderPath = "log\\";
        private const string AppTitle = "SE TIA portal - constructor (V0.4.0)";
        #endregion
        #region Fields
        private readonly IExcelService _excelService;
        private readonly ITiaService _tiaService;
        private readonly IConfigurationService _configuration;
        private readonly ILogger _logger;
        private readonly ProjectExecutor _projectExecutor;
        private string _message;
        private ObservableCollection<string> _logMessages;
        private ProjectState _state = ProjectState.Initial;
        private bool _createTags = true;
        private bool _createInsDB = true;
        private bool _createFC = true;
        private bool _visibleTia = true;
        private bool _closeProject = false;
        private bool _compileProject = true;
        private bool _saveProject = false;
        private bool _isCreateLicenseVisible;
        #endregion
        #region Events
        /// <summary>
        /// Event that is raised when a property value changes
        /// </summary>
        public event PropertyChangedEventHandler PropertyChanged;
        #endregion
        #region Properties
        /// <summary>
        /// Gets the application title
        /// </summary>
        public string Title { get; } = AppTitle;

        /// <summary>
        /// Gets or sets the current status message
        /// </summary>
        public string Message
        {
            get => _message;
            set => SetProperty(ref _message, value);
        }

        /// <summary>
        /// Gets or sets the collection of log messages
        /// </summary>
        public ObservableCollection<string> LogMessages
        {
            get => _logMessages;
            set => SetProperty(ref _logMessages, value);
        }

        /// <summary>
        /// Gets or sets the current project state
        /// </summary>
        public ProjectState State
        {
            get => _state;
            set => SetProperty(ref _state, value);
        }

        /// <summary>
        /// Gets or sets whether to create tags during execution
        /// </summary>
        public bool CreateTags
        {
            get => _createTags;
            set => SetProperty(ref _createTags, value);
        }

        /// <summary>
        /// Gets or sets whether to create instance DBs during execution
        /// </summary>
        public bool CreateInsDB
        {
            get => _createInsDB;
            set => SetProperty(ref _createInsDB, value);
        }

        /// <summary>
        /// Gets or sets whether to create function blocks during execution
        /// </summary>
        public bool CreateFC
        {
            get => _createFC;
            set => SetProperty(ref _createFC, value);
        }

        /// <summary>
        /// Gets or sets whether the TIA portal should be visible
        /// </summary>
        public bool VisibleTia
        {
            get => _visibleTia;
            set => SetProperty(ref _visibleTia, value);
        }

        /// <summary>
        /// Gets or sets whether to close the project after execution
        /// </summary>
        public bool CloseProject
        {
            get => _closeProject;
            set => SetProperty(ref _closeProject, value);
        }

        /// <summary>
        /// Gets or sets whether to compile the project after execution
        /// </summary>
        public bool CompileProject
        {
            get => _compileProject;
            set => SetProperty(ref _compileProject, value);
        }

        /// <summary>
        /// Gets or sets whether to save the project after execution
        /// </summary>
        public bool SaveProject
        {
            get => _saveProject;
            set => SetProperty(ref _saveProject, value);
        }

        /// <summary>
        /// Gets or sets whether the license creation UI should be visible
        /// </summary>
        public bool isCreateLicenseVisible
        {
            get => _isCreateLicenseVisible;
            set { _logger?.Information($"Setting isCreateLicenseVisible to {value}"); SetProperty(ref _isCreateLicenseVisible, value); }
        }

        /// <summary>
        /// Command for creating a license
        /// </summary>
        public ICommand CreateLicenseCommand { get; }

        /// <summary>
        /// Command for reading Excel data
        /// </summary>
        public ICommand ReadExcelCommand { get; }

        /// <summary>
        /// Command for connecting to an opened TIA project
        /// </summary>
        public ICommand ConnectToOpenedCommand { get; }

        /// <summary>
        /// Command for opening a new TIA project
        /// </summary>
        public ICommand OpenTiaCommand { get; }

        /// <summary>
        /// Command for selecting a library
        /// </summary>
        public ICommand SelectLibraryCommand { get; }

        /// <summary>
        /// Command for executing the project
        /// </summary>
        public ICommand ExecuteCommand { get; }
        #endregion
        #region Constructor
        /// <summary>
        /// Initializes a new instance of the MainWindowViewModel class
        /// </summary>
        /// <param name="configuration">Configuration service</param>
        /// <param name="logger">Logger service</param>
        /// <param name="excelService">Excel service</param>
        /// <param name="tiaService">TIA service</param>
        /// <param name="projectManager">Project manager</param>
        /// <param name="libraryManager">Library manager</param>
        /// <param name="plcHardwareManager">PLC hardware manager</param>
        public MainWindowViewModel(IConfigurationService configuration,
            ILogger logger,
            IExcelService excelService,
            ITiaService tiaService,
            IProjectManager projectManager,
            ILibraryManager libraryManager,
            IPlcHardwareManager plcHardwareManager)
        {
            isCreateLicenseVisible = Debugger.IsAttached;

            _excelService = excelService ?? throw new ArgumentNullException(nameof(excelService));
            _tiaService = tiaService ?? throw new ArgumentNullException(nameof(tiaService));
            _logger = logger ?? throw new ArgumentNullException(nameof(logger));
            _configuration = configuration ?? throw new ArgumentNullException(nameof(configuration));
            _ = projectManager ?? throw new ArgumentNullException(nameof(projectManager));
            _ = libraryManager ?? throw new ArgumentNullException(nameof(libraryManager));
            _ = plcHardwareManager ?? throw new ArgumentNullException(nameof(plcHardwareManager));

            _projectExecutor = new ProjectExecutor(configuration, logger, excelService, tiaService, projectManager, libraryManager, plcHardwareManager);

            _logMessages = new ObservableCollection<string>();

            _excelService.MessageUpdated += OnServiceMessageUpdated;
            _tiaService.MessageUpdated += OnServiceMessageUpdated;

            if (_excelService is ExcelService es)
            {
                es.DataReader.MessageUpdated += OnServiceMessageUpdated;
            }

            CreateLicenseCommand = new RelayCommand(ExecuteCreateLicense, () => true);
            ReadExcelCommand = new RelayCommand(ExecuteReadExcel, () => State == ProjectState.PermissionsChecked);
            ConnectToOpenedCommand = new RelayCommand(ExecuteConnectToOpened, () => State == ProjectState.ExcelRead);
            OpenTiaCommand = new RelayCommand(ExecuteOpenTia, () => State == ProjectState.ExcelRead);
            SelectLibraryCommand = new RelayCommand(ExecuteSelectLibrary, () => State == ProjectState.TiaConnected);
            ExecuteCommand = new RelayCommand(Execute, () => State == ProjectState.LibrarySelected);

            var chekPermission = new PermitManager(_configuration, _logger);
            if (chekPermission.CheckLicense())
            {
                State = ProjectState.PermissionsChecked;
            }
        }
        #endregion
        #region Public Methods
        /// <summary>
        /// Closes the application windows and disposes resources
        /// </summary>
        public void CloseWindows()
        {
            LogAndReport("Closing application and disposing resources");
            _excelService.CloseExcelFile();
            _tiaService.DisposeTia();
        }
        #endregion
        #region Private Helper Methods
        private void OnServiceMessageUpdated(object sender, string message)
        {
            Application.Current.Dispatcher.InvokeAsync(() =>
            {
                _logMessages.Insert(0, message);

                if (message.StartsWith("[Main]") || message.StartsWith("[TIA Main]"))
                {
                    Message = message;
                }
            });
        }

        private void ExecuteCreateLicense()
        {
            var chekPermission = new PermitManager(_configuration, _logger);
            chekPermission.GenerateLicense();
        }

        private void ExecuteReadExcel()
        {
            if (!Directory.Exists(LogFolderPath))
                Directory.CreateDirectory(LogFolderPath);

            LogAndReport($"Starting to read Excel file at {DateTime.Now}");

            if (_excelService.CreateExcelDB())
            {
                State = ProjectState.ExcelRead;
                LogAndReport("Reading Excel file completed. Select needed action for TIA project");
            }
        }

        private void ExecuteConnectToOpened()
        {
            VisibleTia = true;
            LogAndReport("Trying to connect to an opened TIA project");

            if (_projectExecutor.ExecuteToOpenedTiaProject(VisibleTia))
            {
                State = ProjectState.TiaConnected;
                LogAndReport("Connection successful. Please select library path");
            }
        }

        private void ExecuteOpenTia()
        {
            LogAndReport("Trying to open a new TIA project");

            if (_projectExecutor.ExecuteOpenTiaProject(VisibleTia))
            {
                State = ProjectState.TiaConnected;
                LogAndReport("TIA project opened successfully. Please select library path");
            }
        }

        private void ExecuteSelectLibrary()
        {
            LogAndReport("Attempting to select TIA library path");

            if (_projectExecutor.ExecuteSelectLibrary())
            {
                State = ProjectState.LibrarySelected;
                LogAndReport("Library selected successfully. Ready to execute");
            }
        }

        private void Execute()
        {
            State = ProjectState.Executing;
            LogAndReport("Starting project execution");

            _projectExecutor.Execute(CreateTags, CreateInsDB, CreateFC, CloseProject, SaveProject, CompileProject, msg =>
                Application.Current.Dispatcher.InvokeAsync(() =>
                {
                    Message = msg;
                    _logger.Information(msg);
                }));

            State = ProjectState.PermissionsChecked;
            LogAndReport("Project execution completed successfully");
        }

        private void LogAndReport(string message)
        {
            string fullMessage = $"{LogPrefix} {message}";
            Message = fullMessage;
            _logger.Information(fullMessage);
            Application.Current.Dispatcher.InvokeAsync(() =>
            {
                _logMessages.Insert(0, fullMessage);
            });
        }

        private bool SetProperty<T>(ref T field, T value, [CallerMemberName] string propertyName = null)
        {
            if (Equals(field, value)) return false;
            field = value;
            OnPropertyChanged(propertyName);
            return true;
        }

        /// <summary>
        /// Raises the PropertyChanged event
        /// </summary>
        /// <param name="propertyName">Name of the property that changed</param>
        protected void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
        #endregion
    }

    /// <summary>
    /// Implementation of ICommand that relays its functionality to delegates
    /// </summary>
    public class RelayCommand : ICommand
    {
        private readonly Action _execute;
        private readonly Func<bool> _canExecute;

        /// <summary>
        /// Initializes a new instance of the RelayCommand class
        /// </summary>
        /// <param name="execute">The execution logic</param>
        /// <param name="canExecute">The execution status logic</param>
        public RelayCommand(Action execute, Func<bool> canExecute = null)
        {
            _execute = execute;
            _canExecute = canExecute;
        }

        /// <summary>
        /// Event that is raised when the execution status changes
        /// </summary>
        public event EventHandler CanExecuteChanged
        {
            add { CommandManager.RequerySuggested += value; }
            remove { CommandManager.RequerySuggested -= value; }
        }

        /// <summary>
        /// Determines whether the command can execute in its current state
        /// </summary>
        /// <param name="parameter">Data used by the command</param>
        /// <returns>True if the command can execute, otherwise false</returns>
        public bool CanExecute(object parameter)
        {
            return _canExecute == null || _canExecute();
        }

        /// <summary>
        /// Executes the command
        /// </summary>
        /// <param name="parameter">Data used by the command</param>
        public void Execute(object parameter)
        {
            _execute();
        }
    }
}