// Ignore Spelling: Plc Conf

using System;
using System.Linq;
using Siemens.Engineering.SW;
using Siemens.Engineering.SW.Blocks;
using Siemens.Engineering.Library.MasterCopies;
using Serilog;
using Siemens.Engineering.Library;
using Siemens.Engineering.SW.Types;

namespace seConfSW.Services
{
    /// <summary>
    /// Provides functionality to manage Master Copies in the TIA Portal.
    /// Handles copying of blocks and UDTs from Master Copy folders to PLC projects.
    /// </summary>
    public class PlcMasterCopyManager : IPlcMasterCopyManager
    {
        #region Constants
        /// <summary>
        /// Prefix for log messages from this class
        /// </summary>
        private const string LogPrefix = "[TIA/McM]";
        #endregion
        #region Properties

        /// <summary>
        /// Logger instance for recording operations and errors
        /// </summary>
        private readonly ILogger _logger;

        #endregion
        #region Constructor

        /// <summary>
        /// Initializes a new instance of the <see cref="PlcMasterCopyManager"/> class.
        /// </summary>
        /// <param name="logger">The logger instance for recording operations</param>
        /// <exception cref="ArgumentNullException">Thrown when logger is null</exception>
        public PlcMasterCopyManager(ILogger logger)
        {
            _logger = logger ?? throw new ArgumentNullException(nameof(logger));
        }

        #endregion
        #region Public Methods

        /// <inheritdoc/>
        public bool CopyBlocksFromMasterCopyFolder(PlcSoftware plcSoftware, UserGlobalLibrary globalLibrary, string blockName, PlcBlockGroup group = null)
        {
            // Validate input parameters
            if (plcSoftware == null) throw new ArgumentNullException(nameof(plcSoftware));
            if (globalLibrary == null) throw new ArgumentNullException(nameof(globalLibrary));
            if (string.IsNullOrEmpty(blockName)) throw new ArgumentException("Block name cannot be null or empty.", nameof(blockName));

            try
            {
                // Split the block name into hierarchy parts
                var hierarchy = blockName.Split('.');
                if (hierarchy.Length <= 1)
                {
                    _logger.Warning($"{LogPrefix} Block name must contain at least one hierarchy level.");
                    return false;
                }

                // Find the root folder in the global library
                var userFolder = globalLibrary.MasterCopyFolder.Folders.FirstOrDefault(f => f.Name == hierarchy[0]);
                if (userFolder == null)
                {
                    _logger.Warning($"{LogPrefix} Root folder not found in Master Copy library: {hierarchy[0]}");
                    return false;
                }

                // Find the Master Copy in the nested folders
                var masterCopy = FindMasterCopy(userFolder, hierarchy.Skip(1).ToArray());
                if (masterCopy == null)
                {
                    _logger.Warning($"{LogPrefix} Master Copy not found: {blockName}");
                    return false;
                }

                // Determine the target group (system group or user-defined group)
                PlcBlockGroup targetGroup = plcSoftware.BlockGroup;
                if (group != null)
                {
                    targetGroup = group;
                }

                // Check if the block already exists in the target group
                if (targetGroup.Blocks.Find(masterCopy.Name) == null)
                {
                    // Create the block from the Master Copy
                    targetGroup.Blocks.CreateFrom(masterCopy);
                    _logger.Information($"{LogPrefix} Created block from Master Copy: {masterCopy.Name}");
                    return true;
                }

                _logger.Information($"{LogPrefix} Block already exists in the target group: {masterCopy.Name}");
                return false;
            }
            catch (Exception ex)
            {
                _logger.Error(ex, $"{LogPrefix} Error occurred while copying blocks from Master Copy folder.");
                return false;
            }
        }

        /// <inheritdoc/>
        public bool CopyUDTFromMasterCopyFolder(PlcSoftware plcSoftware, UserGlobalLibrary globalLibrary, string blockName, string groupName = null)
        {
            // Validate input parameters
            if (plcSoftware == null) throw new ArgumentNullException(nameof(plcSoftware));
            if (globalLibrary == null) throw new ArgumentNullException(nameof(globalLibrary));
            if (string.IsNullOrEmpty(blockName)) throw new ArgumentException("Block name cannot be null or empty.", nameof(blockName));

            try
            {
                // Split the block name into hierarchy parts
                var hierarchy = blockName.Split('.');
                if (hierarchy.Length <= 1)
                {
                    _logger.Warning($"{LogPrefix} Block name must contain at least one hierarchy level.");
                    return false;
                }

                // Find the root folder in the global library
                var userFolder = globalLibrary.MasterCopyFolder.Folders.FirstOrDefault(f => f.Name == hierarchy[0]);
                if (userFolder == null)
                {
                    _logger.Warning($"{LogPrefix} Root folder not found in Master Copy library: {hierarchy[0]}");
                    return false;
                }

                // Find the Master Copy in the nested folders
                var masterCopy = FindMasterCopy(userFolder, hierarchy.Skip(1).ToArray());
                if (masterCopy == null)
                {
                    _logger.Warning($"{LogPrefix} Master Copy not found: {blockName}");
                    return false;
                }

                // Determine the target group (system group or user-defined group)
                PlcTypeGroup targetGroup = plcSoftware.TypeGroup;
                if (!string.IsNullOrEmpty(groupName))
                {
                    var userGroup = targetGroup.Groups.Find(groupName);
                    if (userGroup == null)
                    {
                        userGroup = targetGroup.Groups.Create(groupName);
                        _logger.Information($"{LogPrefix} Created group for UDT: {groupName}");
                    }
                    targetGroup = userGroup;
                }

                // Check if the UDT already exists in the target group
                if (targetGroup.Types.Find(masterCopy.Name) == null)
                {
                    // Create the UDT from the Master Copy
                    targetGroup.Types.CreateFrom(masterCopy);
                    _logger.Information($"{LogPrefix} Created UDT from Master Copy: {masterCopy.Name}");
                    return true;
                }

                _logger.Information($"{LogPrefix} UDT already exists in the target group: {masterCopy.Name}");
                return false;
            }
            catch (Exception ex)
            {
                _logger.Error(ex, $"{LogPrefix} Error occurred while copying UDT from Master Copy folder.");
                return false;
            }
        }

        #endregion
        #region Private Helper Methods

        /// <summary>
        /// Recursively searches for a Master Copy in nested folders
        /// </summary>
        /// <param name="folder">Root folder to start search from</param>
        /// <param name="hierarchy">Array of folder names representing the hierarchy path</param>
        /// <returns>Found MasterCopy or null if not found</returns>
        private MasterCopy FindMasterCopy(MasterCopyUserFolder folder, string[] hierarchy)
        {
            try
            {
                // Traverse the hierarchy to find the Master Copy
                foreach (var folderName in hierarchy.Take(hierarchy.Length - 1))
                {
                    folder = folder.Folders.FirstOrDefault(f => f.Name == folderName);
                    if (folder == null)
                    {
                        _logger.Warning($"{LogPrefix} Folder not found in hierarchy: {folderName}");
                        return null;
                    }
                }

                // Find the Master Copy in the final folder
                var masterCopy = folder.MasterCopies.FirstOrDefault(mc => mc.Name.Equals(hierarchy.Last(), StringComparison.OrdinalIgnoreCase));
                return masterCopy;
            }
            catch (Exception ex)
            {
                _logger.Error(ex, $"{LogPrefix} Error occurred while searching for Master Copy.");
                return null;
            }
        }

        #endregion
    }
}