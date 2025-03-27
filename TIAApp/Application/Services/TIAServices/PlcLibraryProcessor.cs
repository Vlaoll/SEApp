// Ignore Spelling: Eq Plc Conf Prj

using Microsoft.Extensions.DependencyInjection;
using seConfSW.Domain.Models;
using Serilog;
using Siemens.Engineering.Library;
using Siemens.Engineering.SW;
using Siemens.Engineering.SW.Blocks;
using System;
using System.IO;

namespace seConfSW.Services
{
    /// <summary>
    /// Service for processing PLC library operations including updating project libraries,
    /// support blocks, and type blocks from global libraries.
    /// </summary>
    public class PlcLibraryProcessor
    {
        #region Constants
        /// <summary>
        /// Prefix for log messages from this service
        /// </summary>
        private const string LogPrefix = "[TIA/Library]";
        #endregion
        #region Properties
        /// <summary>
        /// Event raised when a message needs to be updated in the UI
        /// </summary>
        public event EventHandler<string> MessageUpdated;
        #endregion
        #region Private Fields
        private readonly IConfigurationService _configuration;
        private readonly ILogger _logger;
        private readonly IProjectManager _projectManager;
        private readonly IHierarchyManager _hierarchyManager;
        private readonly IPlcMasterCopyManager _plcMasterCopyManager;
        private readonly ITagManager _tagManager;
        private readonly ILibraryManager _libraryManager;
        #endregion
        #region Constructor
        /// <summary>
        /// Initializes a new instance of the PlcLibraryProcessor class.
        /// Retrieves required services from the dependency injection container.
        /// </summary>
        /// <exception cref="ArgumentNullException">Thrown when any required service is not available</exception>
        public PlcLibraryProcessor()
        {
            _logger = App.ServiceProvider.GetService<ILogger>() ?? throw new ArgumentNullException(nameof(_logger));
            _configuration = App.ServiceProvider.GetService<IConfigurationService>() ?? throw new ArgumentNullException(nameof(_configuration));
            _tagManager = App.ServiceProvider.GetService<ITagManager>() ?? throw new ArgumentNullException(nameof(_tagManager));
            _projectManager = App.ServiceProvider.GetService<IProjectManager>() ?? throw new ArgumentNullException(nameof(_projectManager));
            _hierarchyManager = App.ServiceProvider.GetService<IHierarchyManager>() ?? throw new ArgumentNullException(nameof(_hierarchyManager));
            _plcMasterCopyManager = App.ServiceProvider.GetService<IPlcMasterCopyManager>() ?? throw new ArgumentNullException(nameof(_plcMasterCopyManager));
            _libraryManager = App.ServiceProvider.GetService<ILibraryManager>() ?? throw new ArgumentNullException(nameof(_libraryManager));
        }
        #endregion
        #region Public Methods
        /// <summary>
        /// Updates the project library from a global library file
        /// </summary>
        /// <param name="plcSoftware">The PLC software instance</param>
        /// <param name="dataPLC">PLC configuration data</param>
        /// <param name="libraryPath">Path to the global library file</param>
        /// <returns>True if the update was successful, false otherwise</returns>
        public bool UpdatePrjLibraryFromGlobal(PlcSoftware plcSoftware, dataPLC dataPLC, string libraryPath)
        {
            string fileNameLibrary;
            try
            {
                _logger.Information($"{LogPrefix} Starting update of project library for PLC: {dataPLC.namePLC} from {libraryPath}");
                fileNameLibrary = new FileInfo(libraryPath).FullName;
                UserGlobalLibrary globalLibrary = _libraryManager.OpenLibrary(fileNameLibrary);

                _libraryManager.UpdatePrjLibraryFromGlobal(fileNameLibrary, _projectManager.WorkProject.ProjectLibrary);
                _logger.Information($"{LogPrefix} Successfully updated project library for PLC: {dataPLC.namePLC}");
                return true;
            }
            catch (Exception ex)
            {
                _logger.Error($"{LogPrefix} Failed to update project library for PLC {dataPLC.namePLC}: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// Updates support blocks (DBs and UDTs) from a global library
        /// </summary>
        /// <param name="plcSoftware">The PLC software instance</param>
        /// <param name="dataPLC">PLC configuration data</param>
        /// <param name="libraryPath">Path to the global library file</param>
        /// <returns>True if the update was successful, false otherwise</returns>
        public bool UpdateSupportBlocks(PlcSoftware plcSoftware, dataPLC dataPLC, string libraryPath)
        {
            string fileNameLibrary;
            try
            {
                _logger.Information($"{LogPrefix} Starting update of support blocks for PLC: {dataPLC.namePLC} from library {libraryPath}");
                fileNameLibrary = new FileInfo(libraryPath).FullName;

                UserGlobalLibrary globalLibrary = _libraryManager.OpenLibrary(fileNameLibrary);
                foreach (dataSupportBD item in dataPLC.dataSupportBD)
                {
                    try
                    {
                        if (item.type == "DB" && !item.isType)
                        {
                            _logger.Information($"{LogPrefix} Copying DB block: {item.name}");
                            PlcBlockGroup group = _hierarchyManager.CreateBlockGroup(plcSoftware, item.group);
                            _plcMasterCopyManager.CopyBlocksFromMasterCopyFolder(plcSoftware, globalLibrary, item.path + "." + item.name, group);
                            continue;
                        }
                        else if (item.type == "UDT" && !item.isType)
                        {
                            _logger.Information($"{LogPrefix} Copying UDT: {item.name}");
                            _plcMasterCopyManager.CopyUDTFromMasterCopyFolder(plcSoftware, globalLibrary, item.path + "." + item.name, item.group);
                            continue;
                        }
                        else if (item.type == "UDT" && item.isType)
                        {
                            _logger.Information($"{LogPrefix} Generating UDT: {item.name}");
                            var projectLibrary = _projectManager.WorkProject.ProjectLibrary;
                            _libraryManager.GenerateUDTFromLibrary(plcSoftware, projectLibrary, "Common", item.name, item.group);
                        }
                    }
                    catch (Exception ex)
                    {
                        _logger.Error($"{LogPrefix} Failed to update support block {item.name} for PLC {dataPLC.namePLC}: {ex.Message}");
                        return false;
                    }
                }
                _logger.Information($"{LogPrefix} Successfully updated support blocks for PLC: {dataPLC.namePLC}");
                return true;
            }
            catch (Exception ex)
            {
                _logger.Error($"{LogPrefix} Failed to update support blocks for PLC {dataPLC.namePLC}: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// Updates type blocks (FBs) from a global library
        /// </summary>
        /// <param name="plcSoftware">The PLC software instance</param>
        /// <param name="dataPLC">PLC configuration data</param>
        /// <param name="libraryPath">Path to the global library file</param>
        /// <returns>True if the update was successful, false otherwise</returns>
        public bool UpdateTypeBlocks(PlcSoftware plcSoftware, dataPLC dataPLC, string libraryPath)
        {
            string fileNameLibrary;
            try
            {
                _logger.Information($"{LogPrefix} Starting update of type blocks for PLC: {dataPLC.namePLC} from library {libraryPath}");
                fileNameLibrary = new FileInfo(libraryPath).FullName;

                UserGlobalLibrary globalLibrary = _libraryManager.OpenLibrary(fileNameLibrary);

                foreach (dataEq equipment in dataPLC.Equipment)
                {
                    foreach (dataLibrary block in equipment.FB)
                    {
                        try
                        {
                            if (!block.isType)
                            {
                                _logger.Information($"{LogPrefix} Copying block: {block.name}");
                                PlcBlockGroup group = _hierarchyManager.CreateBlockGroup(plcSoftware, block.group);
                                _plcMasterCopyManager.CopyBlocksFromMasterCopyFolder(plcSoftware, globalLibrary, block.path + "." + block.name, group);
                            }
                            else
                            {
                                _logger.Information($"{LogPrefix} Generating block: {block.name}");
                                var projectLibrary = _projectManager.WorkProject.ProjectLibrary;
                                _libraryManager.GenerateBlockFromLibrary(plcSoftware, projectLibrary, block.path, block.name, block.group);
                            }
                        }
                        catch (Exception ex)
                        {
                            _logger.Error($"{LogPrefix} Failed to update type block {block.name} for PLC {dataPLC.namePLC}: {ex.Message}");
                            return false;
                        }
                    }
                }
                _logger.Information("{LogPrefix} Closing library");
                _libraryManager.CloseLibrary();
                _logger.Information($"{LogPrefix} Successfully updated type blocks for PLC: {dataPLC.namePLC}");
                return true;
            }
            catch (Exception ex)
            {
                _logger.Error($"{LogPrefix} Failed to update type blocks for PLC {dataPLC.namePLC}: {ex.Message}");
                return false;
            }
        }
        #endregion
    }
}