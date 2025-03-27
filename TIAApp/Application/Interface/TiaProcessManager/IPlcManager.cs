// Ignore Spelling: Plc Conf

using System.Collections.Generic;
using Siemens.Engineering;
using Siemens.Engineering.SW;
using Siemens.Engineering.SW.Blocks;
using Siemens.Engineering.SW.ExternalSources;
using Siemens.Engineering.Library;
using seConfSW.Domain.Models;

namespace seConfSW.Services
{
    /// <summary>
    /// Defines the interface for managing PLC Master Copies in TIA Portal
    /// </summary>
    public interface IPlcMasterCopyManager
    {
        /// <summary>
        /// Copies blocks from Master Copy folder to the specified PLC software
        /// </summary>
        /// <param name="plcSoftware">Target PLC software</param>
        /// <param name="globalLibrary">Source global library containing Master Copies</param>
        /// <param name="blockName">Name of the block to copy (with folder hierarchy)</param>
        /// <param name="group">Optional target block group (uses system group if null)</param>
        /// <returns>True if operation succeeded, false otherwise</returns>
        bool CopyBlocksFromMasterCopyFolder(PlcSoftware plcSoftware, UserGlobalLibrary globalLibrary, string blockName, PlcBlockGroup group = null);

        /// <summary>
        /// Copies UDT (User Defined Type) from Master Copy folder to the specified PLC software
        /// </summary>
        /// <param name="plcSoftware">Target PLC software</param>
        /// <param name="globalLibrary">Source global library containing Master Copies</param>
        /// <param name="blockName">Name of the UDT to copy (with folder hierarchy)</param>
        /// <param name="groupName">Optional target group name (uses system group if null)</param>
        /// <returns>True if operation succeeded, false otherwise</returns>
        bool CopyUDTFromMasterCopyFolder(PlcSoftware plcSoftware, UserGlobalLibrary globalLibrary, string blockName, string groupName = null);
    }

    /// <summary>
    /// Manages operations related to External Sources in the TIA Portal.
    /// </summary>
    public interface IPlcSourceManager
    {
        /// <summary>
        /// Generates source code for a block.
        /// </summary>
        /// <param name="plcSoftware">The PLC software instance.</param>
        /// <param name="blockName">The name of the block.</param>
        /// <param name="path">The path to save the generated source (optional).</param>
        /// <param name="generateOption">The options for generating the source (optional).</param>
        /// <returns>The path to the generated source file, or null if the operation failed.</returns>
        string GenerateSourceBlock(PlcSoftware plcSoftware, string blockName, string path, GenerateOptions generateOption = GenerateOptions.None);

        /// <summary>
        /// Generates source code for a UDT (User-Defined Type).
        /// </summary>
        /// <param name="plcSoftware">The PLC software instance.</param>
        /// <param name="typeName">The name of the UDT.</param>
        /// <param name="path">The path to save the generated source (optional).</param>
        /// <returns>The path to the generated source file, or null if the operation failed.</returns>
        string GenerateSourceUDT(PlcSoftware plcSoftware, string typeName, string path);

        /// <summary>
        /// Clears all external sources in the PLC software.
        /// </summary>
        /// <param name="plcSoftware">The PLC software instance.</param>
        /// <returns>True if the operation was successful; otherwise, false.</returns>
        bool ClearSource(PlcSoftware plcSoftware);

        /// <summary>
        /// Imports a source file into the PLC software.
        /// </summary>
        /// <param name="plcSoftware">The PLC software instance.</param>
        /// <param name="blockName">The name of the block.</param>
        /// <param name="path">The path to the source file (optional).</param>
        /// <param name="type">The file type (optional).</param>
        /// <returns>True if the operation was successful; otherwise, false.</returns>
        bool ImportSource(PlcSoftware plcSoftware, string blockName, string path, string type = null);

        /// <summary>
        /// Exports a block to a file.
        /// </summary>
        /// <param name="plcSoftware">The PLC software instance.</param>
        /// <param name="blockName">The name of the block.</param>
        /// <param name="path">The path to save the exported file (optional).</param>
        /// <returns>True if the operation was successful; otherwise, false.</returns>
        bool ExportBlock(PlcSoftware plcSoftware, string blockName, string path);

        /// <summary>
        /// Imports a block from a file.
        /// </summary>
        /// <param name="plcSoftware">The PLC software instance.</param>
        /// <param name="blockName">The name of the block.</param>
        /// <param name="path">The path to the file (optional).</param>
        /// <param name="groupName">The name of the target group (optional).</param>
        /// <returns>True if the operation was successful; otherwise, false.</returns>
        bool ImportBlock(PlcSoftware plcSoftware, string blockName, string path, string groupName = null);
    }

    /// <summary>
    /// Provides functionality to manage PLC Blocks in the TIA Portal.
    /// This interface extends IPlcSourceManager to include both block management and source management capabilities.
    /// </summary>
    public interface IPlcBlockManager : IPlcSourceManager
    {
        /// <summary>
        /// Creates a new Data Block (DB) in the specified PLC software.
        /// </summary>
        /// <param name="plcSoftware">The PLC software instance where the block will be created</param>
        /// <param name="blockName">The name of the new data block</param>
        /// <param name="number">The block number (use 0 for automatic numbering)</param>
        /// <param name="instanceOfName">The name of the function block this DB will be an instance of</param>
        /// <param name="path">The file system path where source files will be created</param>
        /// <param name="group">The user group where the block should be created</param>
        /// <returns>True if the operation was successful, false otherwise</returns>
        /// <exception cref="ArgumentNullException">Thrown when plcSoftware, group, blockName or instanceOfName is null</exception>
        /// <exception cref="ArgumentException">Thrown when blockName or instanceOfName is empty</exception>
        bool CreateDB(PlcSoftware plcSoftware, string blockName, int number, string instanceOfName, string path, PlcBlockUserGroup group);

        /// <summary>
        /// Creates a new Function Block (FC) with custom content in the specified PLC software.
        /// </summary>
        /// <param name="plcSoftware">The PLC software instance where the block will be created</param>
        /// <param name="blockName">The name of the new function block</param>
        /// <param name="number">The block number (use 0 for automatic numbering)</param>
        /// <param name="blockString">The source code content for the function block</param>
        /// <param name="path">The file system path where source files will be created</param>
        /// <param name="group">The user group where the block should be created</param>
        /// <param name="codeType">The file extension/type for the source file (e.g., ".awl", ".scl")</param>
        /// <returns>True if the operation was successful, false otherwise</returns>
        /// <exception cref="ArgumentNullException">Thrown when plcSoftware, group, blockName, blockString, path or codeType is null</exception>
        /// <exception cref="ArgumentException">Thrown when blockName, blockString, path or codeType is empty</exception>
        bool CreateFC(PlcSoftware plcSoftware, string blockName, int number, string blockString, string path, PlcBlockUserGroup group, string codeType);

        /// <summary>
        /// Creates a new Function Block (FC) from an existing source file in the specified PLC software.
        /// </summary>
        /// <param name="plcSoftware">The PLC software instance where the block will be created</param>
        /// <param name="blockName">The name of the new function block</param>
        /// <param name="number">The block number (use 0 for automatic numbering)</param>
        /// <param name="path">The file system path where source files are located</param>
        /// <param name="group">The user group where the block should be created</param>
        /// <param name="codeType">The file extension/type of the source file (e.g., ".awl", ".scl")</param>
        /// <returns>True if the operation was successful, false otherwise</returns>
        /// <exception cref="ArgumentNullException">Thrown when plcSoftware, group, blockName, path or codeType is null</exception>
        /// <exception cref="ArgumentException">Thrown when blockName, path or codeType is empty</exception>
        bool CreateFC(PlcSoftware plcSoftware, string blockName, int number, string path, PlcBlockUserGroup group, string codeType);

        /// <summary>
        /// Creates an instance Data Block from a Function Block in the specified PLC software.
        /// </summary>
        /// <param name="plcSoftware">The PLC software instance where the block will be created</param>
        /// <param name="blockName">The name of the new instance data block</param>
        /// <param name="number">The block number (use 0 for automatic numbering)</param>
        /// <param name="instanceOfName">The name of the function block this DB will be an instance of</param>
        /// <param name="path">The file system path where source files will be created</param>
        /// <param name="group">The optional user group where the block should be created (null for default group)</param>
        /// <returns>True if the operation was successful, false otherwise</returns>
        /// <exception cref="ArgumentNullException">Thrown when plcSoftware, blockName or instanceOfName is null</exception>
        /// <exception cref="ArgumentException">Thrown when blockName or instanceOfName is empty</exception>
        bool CreateInstanceDB(PlcSoftware plcSoftware, string blockName, int number, string instanceOfName, string path, PlcBlockUserGroup group = null);

        /// <summary>
        /// Generates blocks from external sources in the specified PLC software.
        /// </summary>
        /// <param name="plcSoftware">The PLC software instance where blocks will be generated</param>
        /// <param name="group">The optional user group where blocks should be generated (null for all groups)</param>
        /// <returns>True if the operation was successful, false otherwise</returns>
        /// <exception cref="ArgumentNullException">Thrown when plcSoftware is null</exception>
        bool GenerateBlock(PlcSoftware plcSoftware, PlcBlockUserGroup group = null);

        /// <summary>
        /// Deletes a specific block from the PLC software.
        /// </summary>
        /// <param name="plcSoftware">The PLC software instance containing the block</param>
        /// <param name="blockName">The name of the block to delete</param>
        /// <returns>True if the operation was successful, false otherwise</returns>
        /// <exception cref="ArgumentNullException">Thrown when plcSoftware or blockName is null</exception>
        /// <exception cref="ArgumentException">Thrown when blockName is empty</exception>
        bool DeleteBlock(PlcSoftware plcSoftware, string blockName);

        /// <summary>
        /// Changes the number of a specific block in the PLC software.
        /// </summary>
        /// <param name="plcSoftware">The PLC software instance containing the block</param>
        /// <param name="blockName">The name of the block to modify</param>
        /// <param name="number">The new block number</param>
        /// <param name="group">The group containing the block</param>
        /// <returns>True if the operation was successful, false otherwise</returns>
        /// <exception cref="ArgumentNullException">Thrown when plcSoftware, blockName or group is null</exception>
        /// <exception cref="ArgumentException">Thrown when blockName is empty</exception>
        bool ChangeBlockNumber(PlcSoftware plcSoftware, string blockName, int number, PlcBlockGroup group);

        /// <summary>
        /// Creates a list of Function Blocks from the PLC software and populates the provided dataBlock list.
        /// </summary>
        /// <param name="plcSoftware">The PLC software instance to scan for blocks</param>
        /// <param name="dataBlock">The list to populate with block information</param>
        /// <returns>True if the operation was successful, false otherwise</returns>
        /// <exception cref="ArgumentNullException">Thrown when plcSoftware or dataBlock is null</exception>
        bool CreateListFB(PlcSoftware plcSoftware, List<dataBlock> dataBlock);

        /// <summary>
        /// Retrieves a block by its full path name from the PLC software.
        /// </summary>
        /// <param name="plcSoftware">The PLC software instance to search</param>
        /// <param name="fullPathName">The full path name of the block (using '.' as separator)</param>
        /// <returns>The found PlcBlock instance, or null if not found</returns>
        /// <exception cref="ArgumentNullException">Thrown when plcSoftware or fullPathName is null</exception>
        /// <exception cref="ArgumentException">Thrown when fullPathName is empty</exception>
        PlcBlock GetBlock(PlcSoftware plcSoftware, string fullPathName);
    }

    /// <summary>
    /// Interface for managing PLC hardware configuration in TIA Portal
    /// </summary>
    public interface IPlcHardwareManager
    {
        /// <summary>
        /// Adds a new hardware device to the project
        /// </summary>
        /// <param name="project">TIA Portal project instance</param>
        /// <param name="nameDevice">Name of the device to add</param>
        /// <param name="orderNo">Order number (MLFB) of the device</param>
        /// <param name="version">Hardware version</param>
        /// <returns>True if device was added successfully, false otherwise</returns>
        bool AddHW(Project project, string nameDevice, string orderNo, string version);

        /// <summary>
        /// Retrieves PLC software instance by device name
        /// </summary>
        /// <param name="project">TIA Portal project instance</param>
        /// <param name="plcName">Name of the PLC to find</param>
        /// <returns>PlcSoftware instance if found, null otherwise</returns>
        PlcSoftware GetPLC(Project project, string plcName);
    }
}