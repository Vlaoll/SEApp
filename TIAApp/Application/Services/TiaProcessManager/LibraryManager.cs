// Ignore Spelling: Conf Prj plc

using System;
using System.IO;
using System.Linq;
using Microsoft.Win32;
using Serilog;
using Siemens.Engineering;
using Siemens.Engineering.Library;
using Siemens.Engineering.Library.Types;
using Siemens.Engineering.SW;
using Siemens.Engineering.SW.Blocks;
using Siemens.Engineering.SW.Types;

namespace seConfSW.Services
{
    /// <summary>
    /// Manages operations related to TIA Portal libraries, including opening, updating, and generating blocks.
    /// Implements the ILibraryManager interface.
    /// </summary>
    public class LibraryManager : ILibraryManager
    {
        #region Constants
        private const string LogPrefix = "[TIA/LM]";
        #endregion
        #region Fields
        private readonly ILogger _logger;
        private string _projectLibPath = null;
        private UserGlobalLibrary _globalLibrary = null;
        #endregion
        #region Properties
        /// <inheritdoc />
        public string ProjectLibPath => _projectLibPath;

        /// <inheritdoc />
        public TiaPortal WorkTiaPortal { get; set; }
        #endregion
        #region Constructor
        /// <summary>
        /// Initializes a new instance of the <see cref="LibraryManager"/> class.
        /// </summary>               
        /// <param name="logger">The logger instance.</param>
        /// <exception cref="ArgumentNullException">Thrown if logger is null.</exception>
        public LibraryManager(ILogger logger)
        {
            _logger = logger ?? throw new ArgumentNullException(nameof(logger));
        }
        #endregion
        #region Public Methods
        /// <inheritdoc />
        public UserGlobalLibrary OpenLibrary(string libraryPath)
        {
            if (string.IsNullOrWhiteSpace(libraryPath))
                throw new ArgumentNullException(nameof(libraryPath));

            if (WorkTiaPortal == null)
            {
                _logger.Warning($"{LogPrefix} TIA Portal is not initialized. Cannot search for library.");
                throw new InvalidOperationException("TIA Portal is not initialized. Call ConnectTIA or StartTIA first.");
            }

            try
            {
                var pathLibrari = new FileInfo(libraryPath);
                if (!WorkTiaPortal.GlobalLibraries.Any(g => g?.Path?.FullName == libraryPath))
                {
                    _globalLibrary = WorkTiaPortal.GlobalLibraries.Open(pathLibrari, OpenMode.ReadWrite);
                    _logger.Information($"{LogPrefix} Open global library: {libraryPath}");
                }
                else _globalLibrary = WorkTiaPortal.GlobalLibraries.FirstOrDefault(g => g?.Path?.FullName == libraryPath) as UserGlobalLibrary;
                return _globalLibrary;
            }
            catch (Exception ex)
            {
                _logger.Error($"{LogPrefix} Error while opening global library: {ex.Message}");
                throw;
            }
        }

        /// <inheritdoc />
        public string SearchLibrary(string filter)
        {
            if ((WorkTiaPortal == null) || (WorkTiaPortal?.GetCurrentProcess() == null))
            {
                _logger.Warning($"{LogPrefix} TIA Portal is not initialized. Cannot search for library.");
                _projectLibPath = string.Empty;
                return _projectLibPath;
            }
            if (string.IsNullOrWhiteSpace(filter))
            {
                _logger.Error($"{LogPrefix} Filter is null or empty.");
                _projectLibPath = string.Empty;
                return _projectLibPath;
            }

            var fileSearch = new OpenFileDialog
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
                _projectLibPath = fileSearch.FileName;
                _logger.Information($"{LogPrefix} Selected library: {_projectLibPath}");
                return _projectLibPath;
            }

            _projectLibPath = string.Empty;
            return _projectLibPath;
        }

        /// <inheritdoc />
        public bool CloseLibrary()
        {
            try
            {
                if (_globalLibrary != null)
                {
                    _globalLibrary.Close();
                    _logger.Information($"{LogPrefix} Global library is closed.");
                    return true;
                }
                return false;
            }
            catch (Exception ex)
            {
                _logger.Error($"{LogPrefix} Error while closing library: {ex.Message}");
                throw;
            }
        }

        /// <inheritdoc />
        public bool UpdatePrjLibraryFromGlobal(string libraryPath, ProjectLibrary projectLibrary)
        {
            if (string.IsNullOrWhiteSpace(libraryPath))
                throw new ArgumentNullException(nameof(libraryPath));

            if (WorkTiaPortal == null)
            {
                _logger.Warning($"{LogPrefix} TIA Portal is not initialized. Cannot search for library.");
                throw new InvalidOperationException("TIA Portal is not initialized. Call ConnectTIA or StartTIA first.");
            }

            try
            {
                _logger.Information($"{LogPrefix} Updating project library.");
                var systemTypeFolder = new[] { _globalLibrary.TypeFolder };
                _globalLibrary.UpdateLibrary(systemTypeFolder, projectLibrary);
                _logger.Information($"{LogPrefix} Project Library is updated.");
                return true;
            }
            catch (Exception ex)
            {
                _logger.Error($"{LogPrefix} Error while updating library: {ex.Message}");
                return false;
            }
        }

        /// <inheritdoc />
        public bool GenerateBlockFromLibrary(PlcSoftware plcSoftware, ProjectLibrary projectLibrary, string pathName, string typeName, string groupName = null)
        {
            if (plcSoftware == null) throw new ArgumentNullException(nameof(plcSoftware));
            if (string.IsNullOrWhiteSpace(pathName)) throw new ArgumentNullException(nameof(pathName));
            if (string.IsNullOrWhiteSpace(typeName)) throw new ArgumentNullException(nameof(typeName));

            var group = GetTypeGroup(projectLibrary.TypeFolder, pathName);

            if (group == null)
            {
                _logger.Warning($"{LogPrefix} Group not found for path: {pathName}");
                return false;
            }

            var typeCode = group.Types.Find(typeName)?.Versions.FirstOrDefault(version => version.IsDefault);
            if (typeCode == null)
            {
                _logger.Warning($"{LogPrefix} Type not found: {typeName}");
                return false;
            }

            try
            {
                switch (typeCode)
                {
                    case CodeBlockLibraryTypeVersion typeBlock:
                        var groupBlock = CreateBlockGroup(plcSoftware, groupName);
                        if (groupBlock != null)
                        {
                            CreateBlockFromType(plcSoftware, typeBlock, typeName, groupBlock);
                        }
                        break;

                    case PlcTypeLibraryTypeVersion typePlcType:
                        var groupType = CreateTypeGroup(plcSoftware, groupName);
                        if (groupType != null)
                        {
                            CreateTypeFromLibrary(plcSoftware, typePlcType, typeName, groupType);
                        }
                        break;
                }
                return true;
            }
            catch (Exception ex)
            {
                _logger.Error($"{LogPrefix} Error generating block from library: {ex.Message}");
                return false;
            }
        }

        /// <inheritdoc />
        public bool GenerateUDTFromLibrary(PlcSoftware plcSoftware, ProjectLibrary projectLibrary, string typeName, string blockName, string groupName = null)
        {
            if (plcSoftware == null) throw new ArgumentNullException(nameof(plcSoftware));
            if (string.IsNullOrWhiteSpace(typeName)) throw new ArgumentNullException(nameof(typeName));
            if (string.IsNullOrWhiteSpace(blockName)) throw new ArgumentNullException(nameof(blockName));

            var typeCodeBlock = projectLibrary.TypeFolder.Folders.Find(typeName)
                                  ?.Folders.Find("UDTs")
                                  ?.Types.Find(blockName)
                                  ?.Versions.FirstOrDefault(version => version.IsDefault) as PlcTypeLibraryTypeVersion;

            if (typeCodeBlock == null)
            {
                _logger.Warning($"{LogPrefix} UDT not found: {blockName}");
                return false;
            }

            try
            {
                var blockGroup = plcSoftware.TypeGroup;
                if (groupName != null)
                {
                    var myCreatedGroup = blockGroup.Groups.Find(groupName) ?? blockGroup.Groups.Create(groupName);
                    _logger.Information($"{LogPrefix} Created group for UDT: {groupName}");

                    if (myCreatedGroup.Types.Find(blockName) == null)
                    {
                        myCreatedGroup.Types.CreateFrom(typeCodeBlock);
                        _logger.Information($"{LogPrefix} Imported UDT: {blockName}");
                    }
                }
                else
                {
                    if (blockGroup.Types.Find(blockName) == null)
                    {
                        blockGroup.Types.CreateFrom(typeCodeBlock);
                        _logger.Information($"{LogPrefix} Imported UDT: {blockName}");
                    }
                }
                return true;
            }
            catch (Exception ex)
            {
                _logger.Error($"{LogPrefix} Error generating UDT from library: {ex.Message}");
                return false;
            }
        }

        /// <inheritdoc />
        public bool CleanUpMainLibrary(ProjectLibrary projectLibrary)
        {
            try
            {
                _logger.Information($"{LogPrefix} Cleaning up main library.");
                projectLibrary.CleanUpLibrary(projectLibrary.TypeFolder.Folders, CleanUpMode.DeleteUnusedTypes);
                return true;
            }
            catch (Exception ex)
            {
                _logger.Error($"{LogPrefix} Error during library cleanup: {ex.Message}");
                return false;
            }
        }
        #endregion
        #region Private Helper Methods
        /// <summary>
        /// Creates a block in the PLC software from a library type.
        /// </summary>
        /// <param name="plcSoftware">Target PLC software instance.</param>
        /// <param name="typeBlock">Source block type from library.</param>
        /// <param name="typeName">Name of the type.</param>
        /// <param name="group">Target group for the new block.</param>
        private void CreateBlockFromType(PlcSoftware plcSoftware, CodeBlockLibraryTypeVersion typeBlock, string typeName, PlcBlockUserGroup group)
        {
            if (group != null)
            {
                group.Blocks.CreateFrom(typeBlock);
            }
            else
            {
                plcSoftware.BlockGroup.Blocks.CreateFrom(typeBlock);
            }
            _logger.Information($"{LogPrefix} Imported block: {typeName}");
        }

        /// <summary>
        /// Creates a type in the PLC software from a library type.
        /// </summary>
        /// <param name="plcSoftware">Target PLC software instance.</param>
        /// <param name="typePlcType">Source type from library.</param>
        /// <param name="typeName">Name of the type.</param>
        /// <param name="group">Target group for the new type.</param>
        private void CreateTypeFromLibrary(PlcSoftware plcSoftware, PlcTypeLibraryTypeVersion typePlcType, string typeName, PlcTypeUserGroup group)
        {
            if (group != null)
            {
                group.Types.CreateFrom(typePlcType);
            }
            else
            {
                plcSoftware.TypeGroup.Types.CreateFrom(typePlcType);
            }
            _logger.Information($"{LogPrefix} Imported type: {typeName}");
        }

        /// <summary>
        /// Gets a type group from the library folder hierarchy.
        /// </summary>
        /// <param name="typeFolder">Root type folder.</param>
        /// <param name="typeName">Full path name of the type (using '.' as separator).</param>
        /// <returns>The found LibraryTypeUserFolder or null if not found.</returns>
        private LibraryTypeUserFolder GetTypeGroup(LibraryTypeSystemFolder typeFolder, string typeName)
        {
            if (typeFolder == null)
                throw new ArgumentNullException(nameof(typeFolder));
            if (string.IsNullOrWhiteSpace(typeName))
                throw new ArgumentNullException(nameof(typeName));

            var hierarchy = typeName.Split('.');
            LibraryTypeUserFolder group = null;

            foreach (var folderName in hierarchy)
            {
                group = group == null ? typeFolder.Folders.Find(folderName) : group.Folders.Find(folderName);
                if (group == null)
                {
                    _logger.Warning($"{LogPrefix} Wrong path or missing folder: {folderName}");
                    return null;
                }
            }

            return group;
        }

        /// <summary>
        /// Creates a type group in the PLC software.
        /// </summary>
        /// <param name="plcSoftware">Target PLC software instance.</param>
        /// <param name="groupName">Name of the group to create (can be hierarchical using '.').</param>
        /// <returns>The created or existing PlcTypeUserGroup.</returns>
        private PlcTypeUserGroup CreateTypeGroup(PlcSoftware plcSoftware, string groupName)
        {
            if (plcSoftware == null) throw new ArgumentNullException(nameof(plcSoftware));
            if (string.IsNullOrEmpty(groupName)) throw new ArgumentException("Group name cannot be null or empty", nameof(groupName));

            try
            {
                PlcTypeSystemGroup systemGroup = plcSoftware.TypeGroup;
                PlcTypeUserGroupComposition groupComposition = systemGroup.Groups;
                PlcTypeUserGroup myCreatedGroup = null;

                foreach (string currentGroupName in groupName.Split('.'))
                {
                    myCreatedGroup = groupComposition.Find(currentGroupName);
                    if (myCreatedGroup == null)
                    {
                        myCreatedGroup = groupComposition.Create(currentGroupName);
                        _logger.Information($"{LogPrefix} Create group for types: {currentGroupName}");
                    }

                    groupComposition = myCreatedGroup.Groups;
                }

                return myCreatedGroup;
            }
            catch (Exception ex)
            {
                _logger.Error($"{LogPrefix} Error creating type group: {ex.Message}");
                return null;
            }
        }

        /// <summary>
        /// Creates a block group in the PLC software.
        /// </summary>
        /// <param name="plcSoftware">Target PLC software instance.</param>
        /// <param name="groupName">Name of the group to create (can be hierarchical using '.').</param>
        /// <returns>The created or existing PlcBlockUserGroup.</returns>
        private PlcBlockUserGroup CreateBlockGroup(PlcSoftware plcSoftware, string groupName)
        {
            if (plcSoftware == null) throw new ArgumentNullException(nameof(plcSoftware));
            if (string.IsNullOrEmpty(groupName)) throw new ArgumentException("Group name cannot be null or empty", nameof(groupName));

            try
            {
                PlcBlockSystemGroup systemGroup = plcSoftware.BlockGroup;
                PlcBlockUserGroupComposition groupComposition = systemGroup.Groups;
                PlcBlockUserGroup myCreatedGroup = null;

                foreach (string currentGroupName in groupName.Split('.'))
                {
                    myCreatedGroup = groupComposition.Find(currentGroupName);

                    if (myCreatedGroup == null)
                    {
                        myCreatedGroup = groupComposition.Create(currentGroupName);
                        _logger.Information($"{LogPrefix} Create group for blocks: {currentGroupName}");
                    }
                    groupComposition = myCreatedGroup.Groups;
                }

                return myCreatedGroup;
            }
            catch (Exception ex)
            {
                _logger.Error($"{LogPrefix} Error creating block group: {ex.Message}");
                return null;
            }
        }
        #endregion
    }
}