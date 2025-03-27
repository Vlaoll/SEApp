// Ignore Spelling: plc Conf Eq Prj
using System;
using System.Collections.Generic;
using System.IO;
using Microsoft.Win32;
using Serilog;
using Siemens.Engineering;

namespace seConfSW.Services
{
    /// <summary>
    /// Manages TIA Portal projects, including opening, creating, saving, and closing projects.
    /// </summary>
    public class ProjectManager : IProjectManager
    {
        #region Constants
        private const string LogPrefix = "[TIA/PM]";
        private readonly ILogger _logger;
        private TiaPortalProcess _tiaProcess;
        private string _projectPath = null;
        #endregion
        #region Properties

        /// <inheritdoc />
        public string ProjectPath => _projectPath;

        /// <inheritdoc />
        public TiaPortal WorkTiaPortal { get; private set; }

        /// <inheritdoc />
        public Project WorkProject { get; private set; }

        #endregion
        #region Constructor

        /// <summary>
        /// Initializes a new instance of the <see cref="ProjectManager"/> class.
        /// </summary>
        /// <param name="logger">The logger instance.</param>
        /// <exception cref="ArgumentNullException">Thrown if <paramref name="logger"/> is null.</exception>
        public ProjectManager(ILogger logger)
        {
            _logger = logger ?? throw new ArgumentNullException(nameof(logger));
        }

        #endregion
        #region Public Methods

        /// <inheritdoc />
        public bool StartTIA(bool isVisibleTia = false)
        {
            if (isVisibleTia)
            {
                WorkTiaPortal = new TiaPortal(TiaPortalMode.WithUserInterface);
                _tiaProcess = TiaPortal.GetProcesses()[0];
                _logger.Information($"{LogPrefix} TIA Portal started with user interface");
            }
            else
            {
                WorkTiaPortal = new TiaPortal(TiaPortalMode.WithoutUserInterface);
                _logger.Information($"{LogPrefix} TIA Portal started without user interface");
            }
            return WorkTiaPortal != null;
        }

        /// <inheritdoc />
        public ConnectionStatus ConnectToTiaPortal()
        {
            if (WorkTiaPortal?.GetCurrentProcess() != null)
            {
                _logger.Information($"{LogPrefix} Already connected to TIA Portal.");
                return ConnectionStatus.ConnectedSuccessfully;
            }

            _logger.Information($"{LogPrefix} Trying to connect to project");
            IList<TiaPortalProcess> processes = TiaPortal.GetProcesses();

            switch (processes.Count)
            {
                case 0:
                    _logger.Information($"{LogPrefix} No running instance of TIA Portal was found!");
                    return ConnectionStatus.NoInstanceFound;
                case 1:
                    _tiaProcess = processes[0];
                    WorkTiaPortal = _tiaProcess.Attach();

                    if (WorkTiaPortal.Projects.Count <= 0)
                    {
                        _logger.Information($"{LogPrefix} No TIA Portal Project was found!");
                        return ConnectionStatus.NoProjectFound;
                    }

                    WorkProject = WorkTiaPortal.Projects[0];
                    _logger.Information($"{LogPrefix} Connected to Project: {WorkTiaPortal.Projects[0].Name}");
                    return ConnectionStatus.ConnectedSuccessfully;
                default:
                    _logger.Information($"{LogPrefix} More than one running instance of TIA Portal was found!");
                    return ConnectionStatus.MultipleInstancesFound;
            }
        }

        /// <inheritdoc />
        public Project OpenProject(TiaPortal tiaPortal, string projectPath)
        {
            if (tiaPortal == null)
            {
                throw new InvalidOperationException("TIA Portal is not initialized. Call ConnectTIA or StartTIA first.");
            }

            if (string.IsNullOrWhiteSpace(projectPath))
            {
                _logger.Error($"{LogPrefix} Project path is null or empty.");
                return null;
            }

            try
            {
                WorkProject = tiaPortal.Projects.Open(new FileInfo(projectPath));
                _projectPath = projectPath;
                _logger.Information($"{LogPrefix} Project opened: {WorkProject.Name}");
                return WorkProject;
            }
            catch (Exception ex)
            {
                _logger.Error($"{LogPrefix} Error opening project: {ex.Message}");
                return null;
            }
        }

        /// <inheritdoc />
        public Project CreateProject(TiaPortal tiaPortal, string projectPath)
        {
            if (string.IsNullOrWhiteSpace(projectPath))
            {
                _logger.Error($"{LogPrefix} Project path is null or empty.");
                return null;
            }

            try
            {
                DirectoryInfo directory = new DirectoryInfo(Path.GetDirectoryName(projectPath));
                WorkProject = tiaPortal.Projects.Create(directory, Path.GetFileNameWithoutExtension(projectPath));
                _projectPath = projectPath;
                _logger.Information($"{LogPrefix} Project created: {WorkProject.Name}");
                return WorkProject;
            }
            catch (Exception ex)
            {
                _logger.Error($"{LogPrefix} Error creating project: {ex.Message}");
                return null;
            }
        }

        /// <inheritdoc />
        public void SaveProject()
        {
            try
            {
                if (WorkProject != null)
                {
                    WorkProject.Save();
                    _logger.Information($"{LogPrefix} Project saved: {WorkProject.Name}");
                }
            }
            catch (Exception ex)
            {
                _logger.Error($"{LogPrefix} Error saving project: {ex.Message}");
            }
        }

        /// <inheritdoc />
        public void CloseProject()
        {
            try
            {
                if (WorkProject != null)
                {
                    WorkProject.Close();
                    _logger.Information($"{LogPrefix} Project closed: {WorkProject.Name}");
                    WorkProject = null;
                    _projectPath = null;
                }
            }
            catch (Exception ex)
            {
                _logger.Error($"{LogPrefix} Error closing project: {ex.Message}");
            }
        }

        /// <inheritdoc />
        public string SearchProject(string filter)
        {
            if (string.IsNullOrWhiteSpace(filter))
            {
                _logger.Error($"{LogPrefix} Filter is null or empty.");
                _projectPath = string.Empty;
                return _projectPath;
            }

            OpenFileDialog fileSearch = new OpenFileDialog
            {
                Multiselect = false,
                ValidateNames = true,
                DereferenceLinks = false,
                Filter = filter,
                RestoreDirectory = true,
                InitialDirectory = Environment.CurrentDirectory
            };

            if (fileSearch.ShowDialog() == true)
            {
                _projectPath = fileSearch.FileName;
                _logger.Information($"{LogPrefix} Selected project: {_projectPath}");
                return _projectPath;
            }

            _projectPath = string.Empty;
            return _projectPath;
        }

        /// <inheritdoc />
        public void Dispose()
        {
            if (WorkTiaPortal != null)
            {
                WorkTiaPortal.Dispose();
                _logger.Information($"{LogPrefix} TIA Portal disposed");
                _tiaProcess?.Dispose();
                WorkTiaPortal = null;
            }
        }

        #endregion
    }
}