// Ignore Spelling: plc Conf Eq Prj

using Siemens.Engineering.Library;
using Siemens.Engineering.SW.Blocks;
using seConfSW.Domain.Models;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using Serilog;
using Siemens.Engineering.SW;

namespace seConfSW.Services
{
    /// <summary>
    /// Service for managing TIA Portal operations including PLC blocks, tags, libraries, and project management.
    /// Implements the ITiaService interface.
    /// </summary>
    public class TiaService : ITiaService
    {
        #region Constants
        /// <summary>
        /// Prefix for log messages from this service
        /// </summary>
        private const string LogPrefix = "[TIA]";
        #endregion
        #region Events

        /// <summary>
        /// Event that is triggered when a message needs to be updated
        /// </summary>
        public event EventHandler<string> MessageUpdated;

        #endregion
        #region Private Fields
        private readonly IConfigurationService _configuration;
        private readonly ILogger _logger;
        private readonly IProjectManager _projectManager;
        private readonly IPlcBlockManager _plcBlockManager;
        private readonly ICompilerManager _compilerProcess;
        private readonly IHierarchyManager _hierarchyManager;
        private readonly IPlcSourceManager _plcSourceManager;
        #endregion
        #region Constructor
        /// <summary>
        /// Initializes a new instance of the TiaService class
        /// </summary>
        /// <param name="logger">Logger instance for logging messages</param>
        /// <param name="configuration">Configuration service for application settings</param>
        /// <param name="projectManager">Manager for TIA project operations</param>
        /// <param name="plcBlockManager">Manager for PLC block operations</param>
        /// <param name="compilerProcess">Manager for compilation processes</param>
        /// <param name="hierarchyManager">Manager for project hierarchy operations</param>
        /// <param name="plcMasterCopyManager">Manager for PLC master copy operations (unused in this service)</param>
        /// <param name="plcSourceManager">Manager for PLC source operations</param>
        /// <param name="plcHardwareManager">Manager for PLC hardware operations (unused in this service)</param>
        /// <param name="tagManager">Manager for tag operations (unused in this service)</param>
        /// <param name="libraryManager">Manager for library operations (unused in this service)</param>
        /// <exception cref="ArgumentNullException">Thrown when any required dependency is null</exception>
        public TiaService(
            ILogger logger,
            IConfigurationService configuration,
            IProjectManager projectManager,
            IPlcBlockManager plcBlockManager,
            ICompilerManager compilerProcess,
            IHierarchyManager hierarchyManager,
            IPlcMasterCopyManager plcMasterCopyManager,
            IPlcSourceManager plcSourceManager,
            IPlcHardwareManager plcHardwareManager,
            ITagManager tagManager,
            ILibraryManager libraryManager)
        {
            _configuration = configuration ?? throw new ArgumentNullException(nameof(configuration));
            _logger = logger ?? throw new ArgumentNullException(nameof(logger));
            _hierarchyManager = hierarchyManager ?? throw new ArgumentNullException(nameof(hierarchyManager));
            _plcBlockManager = plcBlockManager ?? throw new ArgumentNullException(nameof(plcBlockManager));
            _compilerProcess = compilerProcess ?? throw new ArgumentNullException(nameof(compilerProcess));
            _projectManager = projectManager ?? throw new ArgumentNullException(nameof(projectManager));
            _plcSourceManager = plcSourceManager ?? throw new ArgumentNullException(nameof(plcSourceManager));
        }
        #endregion
        #region Public Methods - Data Block Operations
        /// <summary>
        /// Adds a value to a data block in the specified PLC software
        /// </summary>
        /// <param name="plcSoftware">The PLC software instance to modify</param>
        /// <param name="dataPLC">Data containing information about the value to add</param>
        /// <returns>True if the operation was successful, false otherwise</returns>
        public bool AddValueToDataBlock(PlcSoftware plcSoftware, dataPLC dataPLC)
        {
            var processor = new PlcDataBlockProcessor(_configuration.DefaultSourcePath);
            return processor.AddValueToDataBlock(plcSoftware, dataPLC);
        }
        #endregion
        #region Public Methods - Tag Operations
        /// <summary>
        /// Creates common user constants in the specified PLC software
        /// </summary>
        /// <param name="plcSoftware">The PLC software instance to modify</param>
        /// <param name="dataPLC">Data containing information about the constants to create</param>
        /// <returns>True if the operation was successful, false otherwise</returns>
        public bool CreateCommonUserConstants(PlcSoftware plcSoftware, dataPLC dataPLC)
        {
            var processor = new PlcUserTagProcessor();
            return processor.CreateCommonUserConstants(plcSoftware, dataPLC);
        }

        /// <summary>
        /// Creates equipment constants in the specified PLC software
        /// </summary>
        /// <param name="plcSoftware">The PLC software instance to modify</param>
        /// <param name="dataPLC">Data containing information about the constants to create</param>
        /// <returns>True if the operation was successful, false otherwise</returns>
        public bool CreateEqConstants(PlcSoftware plcSoftware, dataPLC dataPLC)
        {
            var processor = new PlcUserTagProcessor();
            return processor.CreateEqConstants(plcSoftware, dataPLC);
        }

        /// <summary>
        /// Creates tags in the PLC from a file
        /// </summary>
        /// <param name="plcSoftware">The PLC software instance to modify</param>
        /// <param name="dataPLC">Data containing information about the tags to create</param>
        /// <returns>True if the operation was successful, false otherwise</returns>
        public bool CreateTagsFromFile(PlcSoftware plcSoftware, dataPLC dataPLC)
        {
            var processor = new PlcUserTagProcessor();
            return processor.CreateTagsFromFile(plcSoftware, dataPLC);
        }
        #endregion
        #region Public Methods - Library Operations
        /// <summary>
        /// Updates the project library from a global library source
        /// </summary>
        /// <param name="plcSoftware">The PLC software instance to modify</param>
        /// <param name="dataPLC">Data containing information about the update</param>
        /// <param name="libraryPath">Path to the global library</param>
        /// <returns>True if the operation was successful, false otherwise</returns>
        public bool UpdatePrjLibraryFromGlobal(PlcSoftware plcSoftware, dataPLC dataPLC, string libraryPath)
        {
            var processor = new PlcLibraryProcessor();
            return processor.UpdatePrjLibraryFromGlobal(plcSoftware, dataPLC, libraryPath);
        }

        /// <summary>
        /// Updates support blocks in the PLC from a library
        /// </summary>
        /// <param name="plcSoftware">The PLC software instance to modify</param>
        /// <param name="dataPLC">Data containing information about the update</param>
        /// <param name="libraryPath">Path to the library containing support blocks</param>
        /// <returns>True if the operation was successful, false otherwise</returns>
        public bool UpdateSupportBlocks(PlcSoftware plcSoftware, dataPLC dataPLC, string libraryPath)
        {
            var processor = new PlcLibraryProcessor();
            return processor.UpdateSupportBlocks(plcSoftware, dataPLC, libraryPath);
        }

        /// <summary>
        /// Updates type blocks in the PLC from a library
        /// </summary>
        /// <param name="plcSoftware">The PLC software instance to modify</param>
        /// <param name="dataPLC">Data containing information about the update</param>
        /// <param name="libraryPath">Path to the library containing type blocks</param>
        /// <returns>True if the operation was successful, false otherwise</returns>
        public bool UpdateTypeBlocks(PlcSoftware plcSoftware, dataPLC dataPLC, string libraryPath)
        {
            var processor = new PlcLibraryProcessor();
            return processor.UpdateTypeBlocks(plcSoftware, dataPLC, libraryPath);
        }
        #endregion
        #region Public Methods - Function Block Operations
        /// <summary>
        /// Creates instance blocks in the specified PLC software
        /// </summary>
        /// <param name="plcSoftware">The PLC software instance to modify</param>
        /// <param name="dataPLC">Data containing information about the blocks to create</param>
        /// <returns>True if the operation was successful, false otherwise</returns>
        public bool CreateInstanceBlocks(PlcSoftware plcSoftware, dataPLC dataPLC)
        {
            var processor = new PlcFunctionBlockProcessor(_logger, _configuration, _projectManager, _plcBlockManager, _compilerProcess,
                _plcSourceManager, _hierarchyManager);
            return processor.CreateInstanceBlocks(plcSoftware, dataPLC);
        }

        /// <summary>
        /// Creates template function blocks from Excel data
        /// </summary>
        /// <param name="plcSoftware">The PLC software instance to modify</param>
        /// <param name="dataPLC">Data containing information about the blocks to create</param>
        /// <returns>True if the operation was successful, false otherwise</returns>
        public bool CreateTemplateFCFromExcel(PlcSoftware plcSoftware, dataPLC dataPLC)
        {
            var processor = new PlcFunctionBlockProcessor(_logger, _configuration, _projectManager, _plcBlockManager, _compilerProcess,
                _plcSourceManager, _hierarchyManager);
            return processor.CreateTemplateFcForExtendedType(plcSoftware, dataPLC);
        }

        /// <summary>
        /// Edits function blocks from Excel data and calls all blocks
        /// </summary>
        /// <param name="plcSoftware">The PLC software instance to modify</param>
        /// <param name="dataPLC">Data containing information about the blocks to edit</param>
        /// <param name="listDataPLC">List of additional PLC data for processing</param>
        /// <param name="closeProject">Whether to close the project after operation</param>
        /// <param name="saveProject">Whether to save the project after operation</param>
        /// <param name="compileProject">Whether to compile the project after operation</param>
        /// <returns>True if the operation was successful, false otherwise</returns>
        public bool EditFCFromExcelCallAllBlocks(PlcSoftware plcSoftware, dataPLC dataPLC, List<dataPLC> listDataPLC,
            bool closeProject = false, bool saveProject = false, bool compileProject = false)
        {
            var processor = new PlcFunctionBlockProcessor(_logger, _configuration, _projectManager, _plcBlockManager, _compilerProcess,
                _plcSourceManager, _hierarchyManager);
            return processor.EditFCFromExcelCallAllBlocks(plcSoftware, dataPLC, listDataPLC, closeProject, saveProject, compileProject);
        }
        #endregion
        #region Public Methods - Project Management
        /// <summary>
        /// Disposes TIA Portal resources and cleans up the project manager
        /// </summary>
        public void DisposeTia()
        {
            try
            {
                _logger.Information($"{LogPrefix} Disposing TIA resources");
                _projectManager.Dispose();
                _logger.Information($"{LogPrefix} TIA resources disposed");
            }
            catch (Exception ex)
            {
                _logger.Error($"{LogPrefix} Failed to dispose TIA resources: {ex.Message}");
            }
        }
        #endregion
        #region Private Helper Methods
        protected virtual void OnMessageUpdated(string message)
        {
            MessageUpdated?.Invoke(this, message);
        }
        #endregion
    }
}