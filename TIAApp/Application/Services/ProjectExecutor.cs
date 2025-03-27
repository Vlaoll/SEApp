// Ignore Spelling: Conf plc

using seConfSW.Services;
using Serilog;
using Siemens.Engineering;
using Siemens.Engineering.SW;
using System;
using System.Collections.Generic;

namespace seConfSW.Presentation.ViewModels
{
    /// <summary>
    /// Handles the execution of TIA Portal project operations including project opening, 
    /// library selection, and PLC data processing from Excel.
    /// </summary>
    public class ProjectExecutor
    {
        #region Constants
        /// <summary>
        /// Prefix for all log messages
        /// </summary>
        private const string LogPrefix = "[TIA Main]";
        #endregion
        #region Properties - Dependencies (Injected Services)
        private readonly IExcelService _excelService;
        private readonly ITiaService _tiaService;
        private readonly ILibraryManager _libraryManager;
        private readonly IProjectManager _projectManager;
        private readonly IConfigurationService _configuration;
        private readonly IPlcHardwareManager _plcHardwareManager;
        private readonly ILogger _logger;
        #endregion
        #region Properties - Paths
        private readonly string _exportFolderPath = null;
        private readonly string _sourceFolderPath = null;
        private readonly string _templateFolderPath = null;
        private string _projectPath = null;
        private string _projectLibPath = null;
        #endregion
        #region Constructor
        /// <summary>
        /// Initializes a new instance of the ProjectExecutor class with required services.
        /// </summary>
        /// <param name="configuration">Configuration service</param>
        /// <param name="logger">Logger service</param>
        /// <param name="excelService">Excel data service</param>
        /// <param name="tiaService">TIA Portal operations service</param>
        /// <param name="projectManager">Project management service</param>
        /// <param name="libraryManager">Library management service</param>
        /// <param name="plcHardwareManager">PLC hardware management service</param>
        /// <exception cref="ArgumentNullException">Thrown when any required service is null</exception>
        public ProjectExecutor(
            IConfigurationService configuration,
            ILogger logger,
            IExcelService excelService,
            ITiaService tiaService,
            IProjectManager projectManager,
            ILibraryManager libraryManager,
            IPlcHardwareManager plcHardwareManager)
        {
            _excelService = excelService ?? throw new ArgumentNullException(nameof(excelService));
            _tiaService = tiaService ?? throw new ArgumentNullException(nameof(tiaService));
            _logger = logger ?? throw new ArgumentNullException(nameof(logger));
            _configuration = configuration ?? throw new ArgumentNullException(nameof(configuration));
            _projectManager = projectManager ?? throw new ArgumentNullException(nameof(projectManager));
            _libraryManager = libraryManager ?? throw new ArgumentNullException(nameof(libraryManager));
            _plcHardwareManager = plcHardwareManager ?? throw new ArgumentNullException(nameof(plcHardwareManager));

            _exportFolderPath = _configuration.ExportPath;
            _sourceFolderPath = _configuration.DefaultSourcePath;
            _templateFolderPath = _configuration.TemplatePath;
        }
        #endregion
        #region Public Methods
        /// <summary>
        /// Opens a TIA Portal project with the specified visibility.
        /// </summary>
        /// <param name="isVisible">Whether to make TIA Portal visible</param>
        /// <returns>True if project was successfully opened and connected</returns>
        public bool ExecuteOpenTiaProject(bool isVisible)
        {
            LogAndReport("{LogPrefix} Checking if TIA Portal is available", _logger.Information);
            if (!(_projectManager.ConnectToTiaPortal() == ConnectionStatus.NoInstanceFound))
            {
                LogAndReport("TIA Portal must be closed before opening a new project", _logger.Warning);
                return false;
            }

            _projectPath = string.Empty;
            LogAndReport("Searching for TIA project", _logger.Information);
            _projectPath = _projectManager.SearchProject(_configuration.ProjectFilter);

            if (string.IsNullOrEmpty(_projectPath))
            {
                LogAndReport($"TIA project not selected or invalid path: {_projectPath}", _logger.Warning);
                return false;
            }

            if (!_projectManager.StartTIA(isVisible))
            {
                LogAndReport("TIA Portal didn't start.", _logger.Warning);
                return false;
            }

            if (!(_projectManager.OpenProject(_projectManager.WorkTiaPortal, _projectPath) is Project))
            {
                LogAndReport($"TIA project didn't open at {_projectPath}.", _logger.Warning);
                return false;
            }

            LogAndReport($"Connecting to TIA project at {_projectPath}", _logger.Information);

            if (_projectManager.ConnectToTiaPortal() == ConnectionStatus.ConnectedSuccessfully)
            {
                LogAndReport($"TIA project opened and connected at {_projectPath}", _logger.Information);
                _libraryManager.WorkTiaPortal = _projectManager.WorkTiaPortal;
                return true;
            }

            LogAndReport($"TIA project didn't connect at {_projectPath}", _logger.Warning);
            return false;
        }

        /// <summary>
        /// Connects to an already opened TIA Portal project.
        /// </summary>
        /// <param name="isVisible">Whether to make TIA Portal visible</param>
        /// <returns>True if successfully connected to the opened project</returns>
        public bool ExecuteToOpenedTiaProject(bool isVisible)
        {
            LogAndReport("Attempting to connect to an already opened TIA project", _logger.Information);
            bool iConnected = _projectManager.ConnectToTiaPortal() == ConnectionStatus.ConnectedSuccessfully;
            _libraryManager.WorkTiaPortal = iConnected ? _projectManager.WorkTiaPortal : null;
            LogAndReport(iConnected ? "Successfully connected to opened TIA project" : "Failed to connect to TIA project", _logger.Information);
            return iConnected;
        }

        /// <summary>
        /// Selects a library for the project.
        /// </summary>
        /// <returns>True if library was successfully selected</returns>
        public bool ExecuteSelectLibrary()
        {
            LogAndReport("Searching for library", _logger.Information);
            _projectLibPath = string.Empty;
            _projectLibPath = _libraryManager.SearchLibrary(_configuration.LibraryFilter);

            if (!string.IsNullOrEmpty(_projectLibPath))
            {
                LogAndReport($"Library selected at {_projectLibPath}", _logger.Information);
                return true;
            }

            LogAndReport("Invalid library path", _logger.Warning);
            return false;
        }

        /// <summary>
        /// Executes the main project processing workflow which includes:
        /// - Creating necessary folders
        /// - Processing PLC data from Excel
        /// - Optionally closing/saving/compiling the project
        /// - Providing status updates through callback
        /// </summary>
        /// <param name="createTags">Flag indicating whether to create tags during processing (unused in current implementation)</param>
        /// <param name="createInsDB">Flag indicating whether to create instance DBs during processing (unused in current implementation)</param>
        /// <param name="createFC">Flag indicating whether to create function blocks during processing (unused in current implementation)</param>
        /// <param name="closeProject">Whether to close the project after processing completes</param>
        /// <param name="saveProject">Whether to save the project after processing completes</param>
        /// <param name="compileProject">Whether to compile the project after processing completes</param>
        /// <param name="updateMessage">Callback function for receiving status/progress updates</param>
        
        public void Execute(bool createTags, bool createInsDB, bool createFC, bool closeProject, bool saveProject, bool compileProject, Action<string> updateMessage)
        {
            // Record and log start time for performance tracking
            DateTime startTime = DateTime.Now;
            LogAndReport($"Starting project execution at {startTime}", updateMessage);

            // Ensure all required working directories exist before processing
            // These folders are essential for export, source files, and templates
            Common.CreateNewFolder(_exportFolderPath);     // Folder for exported files
            Common.CreateNewFolder(_sourceFolderPath);     // Folder for source files
            Common.CreateNewFolder(_templateFolderPath);   // Folder for template files

            // Retrieve PLC configuration data from Excel file
            // Each PLC will be processed sequentially
            List<Domain.Models.dataPLC> excelData = _excelService.GetExcelDataReader();

            // Process each PLC configuration from the Excel data
            foreach (var plc in excelData)
            {
                // Process current PLC with the complete dataset
                // This handles all TIA Portal operations for the PLC
                ProcessPlc(plc, excelData, updateMessage, createTags, createInsDB, createFC, closeProject, saveProject, compileProject);
            }

            // Clean up resources if closure is requested
            if (closeProject)
            {
                _tiaService.DisposeTia();        // Properly close TIA Portal instance
                _excelService.CloseExcelFile();  // Close Excel file if open
            }

            // Record and log completion time with duration
            DateTime finishTime = DateTime.Now;
            LogAndReport($"Project execution completed at {finishTime} (Duration: {finishTime - startTime})", updateMessage);
        }

        /// <summary>
        /// Gets the PLC software instance by name.
        /// </summary>
        /// <param name="plcName">Name of the PLC to retrieve</param>
        /// <returns>PlcSoftware instance or null if not found</returns>
        public PlcSoftware GetPLC(string plcName)
        {
            return _plcHardwareManager.GetPLC(_projectManager.WorkProject, plcName);
        }

        /// <summary>
        /// Gets the current project library path.
        /// </summary>
        /// <returns>Path to the project library</returns>
        public string GetProjectLibPath()
        {
            return _libraryManager.ProjectLibPath;
        }
        #endregion
        #region Private Helper Methods
        /// <summary>
        /// Processes a single PLC from the Excel data.
        /// </summary>
        /// <param name="plc">PLC data to process</param>
        /// <param name="excelData">All PLC data from Excel</param>
        /// <param name="updateMessage">Callback for status updates</param>
        /// <param name="closeProject">Whether to close project after processing</param>
        /// <param name="saveProject">Whether to save project after processing</param>
        /// <param name="compileProject">Whether to compile project after processing</param>
        private void ProcessPlc(Domain.Models.dataPLC plc, List<Domain.Models.dataPLC> excelData, Action<string> updateMessage, 
            bool createTags, bool createInsDB, bool createFC, bool closeProject, bool saveProject, bool compileProject)
        {
            LogAndReport($"Starting to process PLC: {plc.namePLC}", updateMessage);

            var plcSoftware = GetPLC(plc.namePLC);

            // Execute all TIA operations for this PLC
            if (createTags)
            {
                _tiaService.AddValueToDataBlock(plcSoftware, plc);
                _tiaService.CreateCommonUserConstants(plcSoftware, plc);
                _tiaService.CreateEqConstants(plcSoftware, plc);
                _tiaService.CreateTagsFromFile(plcSoftware, plc);
            }
            
            var projectLibraryPath = GetProjectLibPath();
            _tiaService.UpdatePrjLibraryFromGlobal(plcSoftware, plc, projectLibraryPath);
            _tiaService.UpdateSupportBlocks(plcSoftware, plc, projectLibraryPath);
            _tiaService.UpdateTypeBlocks(plcSoftware, plc, projectLibraryPath);
            if (createInsDB)
            {
                _tiaService.CreateInstanceBlocks(plcSoftware, plc);
            }
            if (createFC)
            {
                _tiaService.CreateTemplateFCFromExcel(plcSoftware, plc);
                _tiaService.EditFCFromExcelCallAllBlocks(plcSoftware, plc, excelData, closeProject, saveProject, compileProject);
            }               

            LogAndReport($"Finished processing PLC: {plc.namePLC}", updateMessage);
        }

        /// <summary>
        /// Logs a message and reports it through the callback.
        /// </summary>
        /// <param name="message">Message to log and report</param>
        /// <param name="updateMessage">Callback for status updates</param>
        private void LogAndReport(string message, Action<string> updateMessage)
        {
            string fullMessage = $"{LogPrefix} {message}";
            updateMessage(fullMessage);
            _logger.Information(fullMessage);
        }
        #endregion
    }
}