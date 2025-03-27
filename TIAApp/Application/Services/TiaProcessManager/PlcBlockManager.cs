// Ignore Spelling: plc Conf Eq

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Siemens.Engineering.SW;
using Siemens.Engineering.SW.Blocks;
using Serilog;
using seConfSW.Domain.Models;
using Siemens.Engineering.SW.ExternalSources;
using Siemens.Engineering.SW.Types;
using Siemens.Engineering;

namespace seConfSW.Services
{
    /// <summary>
    /// Provides functionality to manage PLC Blocks in the TIA Portal.
    /// </summary>
    public class PlcBlockManager : IPlcBlockManager
    {
        #region Constants
        private const string LogPrefix = "[TIA/BM]";
        #endregion
        #region Properties
        private readonly ILogger _logger;
        #endregion
        #region Constructor
        /// <summary>
        /// Initializes a new instance of the <see cref="PlcBlockManager"/> class.
        /// </summary>
        /// <param name="logger">The logger instance.</param>
        public PlcBlockManager(ILogger logger)
        {
            _logger = logger ?? throw new ArgumentNullException(nameof(logger));
        }
        #endregion
        #region Public Methods (Block Creation)
        /// <inheritdoc />
        public bool CreateDB(PlcSoftware plcSoftware, string blockName, int number, string instanceOfName, string path, PlcBlockUserGroup group)
        {
            if (plcSoftware == null) throw new ArgumentNullException(nameof(plcSoftware));
            if (group == null) throw new ArgumentNullException(nameof(group));
            if (string.IsNullOrEmpty(blockName)) throw new ArgumentException("Block name cannot be null or empty.", nameof(blockName));
            if (string.IsNullOrEmpty(instanceOfName)) throw new ArgumentException("Instance name cannot be null or empty.", nameof(instanceOfName));

            try
            {
                var filename = Path.GetFullPath(Path.Combine(path, instanceOfName + "_instanceDB.db"));
                if (File.Exists(filename))
                {
                    File.Delete(filename);
                    _logger.Information($"{LogPrefix} Deleted existing file: {filename}");
                }

                using (var sw = File.CreateText(filename))
                {
                    sw.WriteLine($"DATA_BLOCK \"{blockName}\"");
                    sw.WriteLine($"\"{instanceOfName}\"");
                    sw.WriteLine("BEGIN");
                    sw.WriteLine("END_DATA_BLOCK");
                }

                _logger.Information($"{LogPrefix} Created instance DB source file: {filename}");

                var rSource = ClearSource(plcSoftware);
                var rImport = ImportSource(plcSoftware, instanceOfName + "_instanceDB", path, ".db");

                if (!rImport)
                {
                    _logger.Warning($"{LogPrefix} Wrong while creating instance DB (import source): {blockName}");
                    return false;
                }

                if (!GenerateBlock(plcSoftware, group))
                {
                    _logger.Warning($"{LogPrefix} Wrong while creating instance DB (creating group): {blockName}");
                    return false;
                }

                if (number != 0 && !ChangeBlockNumber(plcSoftware, blockName, number, group))
                {
                    _logger.Warning($"{LogPrefix} Wrong while creating instance DB (changing number): {blockName}");
                    return false;
                }

                _logger.Information($"{LogPrefix} Created instance DB: {blockName} with number {number}");
                return true;
            }
            catch (Exception ex)
            {
                _logger.Error(ex, $"{LogPrefix} Error occurred while creating instance DB: {blockName}");
                return false;
            }
        }

        /// <inheritdoc />
        public bool CreateFC(PlcSoftware plcSoftware, string blockName, int number, string blockString, string path, PlcBlockUserGroup group, string codeType)
        {
            if (plcSoftware == null) throw new ArgumentNullException(nameof(plcSoftware));
            if (group == null) throw new ArgumentNullException(nameof(group));
            if (string.IsNullOrEmpty(blockName)) throw new ArgumentException("Block name cannot be null or empty.", nameof(blockName));
            if (string.IsNullOrEmpty(blockString)) throw new ArgumentException("Block string cannot be null or empty.", nameof(blockString));
            if (string.IsNullOrEmpty(path)) throw new ArgumentException("Path cannot be null or empty.", nameof(path));
            if (string.IsNullOrEmpty(codeType)) throw new ArgumentException("Code type cannot be null or empty.", nameof(codeType));

            try
            {
                var filename = Path.GetFullPath(Path.Combine(path, blockName + codeType));
                if (File.Exists(filename))
                {
                    File.Delete(filename);
                    _logger.Information($"{LogPrefix} Deleted existing file: {filename}");
                }

                using (var sw = File.CreateText(filename))
                {
                    sw.WriteLine(blockString);
                }

                _logger.Information($"{LogPrefix} Created FC source file: {filename}");

                var rSource = ClearSource(plcSoftware);
                var rImport = ImportSource(plcSoftware, blockName, path, codeType);
                var rDelete = DeleteBlock(plcSoftware, blockName);

                if (!rImport)
                {
                    _logger.Warning($"{LogPrefix} Wrong while creating FC (import source): {blockName}");
                    return false;
                }

                if (!GenerateBlock(plcSoftware, group))
                {
                    _logger.Warning($"{LogPrefix} Wrong while creating FC (creating group): {blockName}");
                    return false;
                }

                if (number != 0 && !ChangeBlockNumber(plcSoftware, blockName, number, group))
                {
                    _logger.Warning($"{LogPrefix} Wrong while creating FC (changing number): {blockName}");
                    return false;
                }

                _logger.Information($"{LogPrefix} Created FC: {blockName} with number {number}");
                return true;
            }
            catch (Exception ex)
            {
                _logger.Error(ex, $"{LogPrefix} Error occurred while creating FC: {blockName}");
                return false;
            }
        }

        /// <inheritdoc />
        public bool CreateFC(PlcSoftware plcSoftware, string blockName, int number, string path, PlcBlockUserGroup group, string codeType)
        {
            if (plcSoftware == null) throw new ArgumentNullException(nameof(plcSoftware));
            if (group == null) throw new ArgumentNullException(nameof(group));
            if (string.IsNullOrEmpty(blockName)) throw new ArgumentException("Block name cannot be null or empty.", nameof(blockName));
            if (string.IsNullOrEmpty(path)) throw new ArgumentException("Path cannot be null or empty.", nameof(path));

            try
            {
                var rSource = ClearSource(plcSoftware);
                var rImport = ImportSource(plcSoftware, blockName, path, codeType);
                var rDelete = DeleteBlock(plcSoftware, blockName);

                if (!rImport)
                {
                    _logger.Warning($"{LogPrefix} Wrong while creating FC (import source): {blockName}");
                    return false;
                }

                if (!GenerateBlock(plcSoftware, group))
                {
                    _logger.Warning($"{LogPrefix} Wrong while creating FC (creating group): {blockName}");
                    return false;
                }

                if (number != 0 && !ChangeBlockNumber(plcSoftware, blockName, number, group))
                {
                    _logger.Warning($"{LogPrefix} Wrong while creating FC (changing number): {blockName}");
                    return false;
                }

                _logger.Information($"{LogPrefix} Created FC: {blockName} with number {number}");
                return true;
            }
            catch (Exception ex)
            {
                _logger.Error(ex, $"{LogPrefix} Error occurred while creating FC: {blockName}");
                return false;
            }
        }

        /// <inheritdoc />
        public bool CreateInstanceDB(PlcSoftware plcSoftware, string blockName, int number, string instanceOfName, string path, PlcBlockUserGroup group = null)
        {
            if (plcSoftware == null) throw new ArgumentNullException(nameof(plcSoftware));
            if (string.IsNullOrEmpty(blockName)) throw new ArgumentException("Block name cannot be null or empty.", nameof(blockName));
            if (string.IsNullOrEmpty(instanceOfName)) throw new ArgumentException("Instance name cannot be null or empty.", nameof(instanceOfName));
            if (string.IsNullOrEmpty(path)) throw new ArgumentException("Path cannot be null or empty.", nameof(path));

            try
            {
                DeleteBlock(plcSoftware, blockName);

                if (group != null)
                {
                    if (group.Blocks.Find(blockName) == null)
                    {
                        group.Blocks.CreateInstanceDB(blockName, number == 0, number == 0 ? 1 : number, instanceOfName);
                        _logger.Information($"{LogPrefix} Created instance DB: {blockName}");
                        return true;
                    }
                }
                else
                {
                    if (plcSoftware.BlockGroup.Blocks.Find(blockName) == null)
                    {
                        plcSoftware.BlockGroup.Blocks.CreateInstanceDB(blockName, false, number, instanceOfName);
                        _logger.Information($"{LogPrefix} Created instance DB: {blockName}");
                        return true;
                    }
                }

                _logger.Warning($"{LogPrefix} Instance DB already exists: {blockName}");
                return false;
            }
            catch (Exception ex)
            {
                _logger.Error(ex, $"{LogPrefix} Error occurred while creating instance DB: {blockName}");
                return false;
            }
        }
        #endregion
        #region Public Methods (Block Generation)
        /// <inheritdoc />
        public bool GenerateBlock(PlcSoftware plcSoftware, PlcBlockUserGroup group = null)
        {
            if (plcSoftware == null) throw new ArgumentNullException(nameof(plcSoftware));

            try
            {
                var sources = plcSoftware.ExternalSourceGroup.ExternalSources;
                if (sources == null || sources.Count <= 0)
                {
                    _logger.Warning($"{LogPrefix} No external sources found to generate blocks.");
                    return false;
                }

                foreach (var externalSource in plcSoftware.ExternalSourceGroup.ExternalSources)
                {
                    if (group != null)
                    {
                        var result = externalSource.GenerateBlocksFromSource(group, GenerateBlockOption.KeepOnError);
                        if (result?.Count > 0)
                        {
                            _logger.Information($"{LogPrefix} Generated blocks from source in group: {group.Name}");
                            return true;
                        }
                        _logger.Warning($"{LogPrefix} Generated blocks from source with error: {group.Name}");
                        return false;
                    }
                    else
                    {
                        var result = externalSource.GenerateBlocksFromSource(GenerateBlockOption.KeepOnError);
                        if (result?.Count > 0)
                        {
                            _logger.Information($"{LogPrefix} Generated blocks from source in system group.");
                            return true;
                        }
                        _logger.Warning($"{LogPrefix} Generated blocks from source with error: {group?.Name}");
                        return false;
                    }
                }
                _logger.Warning($"{LogPrefix} Missing source for generate.");
                return false;
            }
            catch (Exception ex)
            {
                _logger.Error(ex, $"{LogPrefix} Error occurred while generating blocks from external sources.");
                return false;
            }
        }
        #endregion
        #region Public Methods (Block Management)
        /// <inheritdoc />
        public bool DeleteBlock(PlcSoftware plcSoftware, string blockName)
        {
            if (plcSoftware == null) throw new ArgumentNullException(nameof(plcSoftware));
            if (string.IsNullOrEmpty(blockName)) throw new ArgumentException("Block name cannot be null or empty.", nameof(blockName));

            try
            {
                foreach (var group in plcSoftware.BlockGroup.Groups)
                {
                    var block = group.Blocks.Find(blockName);
                    if (block != null)
                    {
                        block.Delete();
                        _logger.Information($"{LogPrefix} Deleted block: {blockName}");
                        return true;
                    }
                }

                _logger.Debug($"{LogPrefix} Block not found: {blockName}");
                return false;
            }
            catch (Exception ex)
            {
                _logger.Error(ex, $"{LogPrefix} Error occurred while deleting block: {blockName}");
                return false;
            }
        }

        /// <inheritdoc />
        public bool ChangeBlockNumber(PlcSoftware plcSoftware, string blockName, int number, PlcBlockGroup group)
        {
            if (plcSoftware == null) throw new ArgumentNullException(nameof(plcSoftware));
            if (string.IsNullOrEmpty(blockName)) throw new ArgumentException("Block name cannot be null or empty.", nameof(blockName));
            if (group == null) throw new ArgumentNullException(nameof(group));

            try
            {
                var block = group.Blocks.Find(blockName);
                if (block != null)
                {
                    block.AutoNumber = false;
                    block.Number = number;
                    _logger.Information($"{LogPrefix} Changed block number: {blockName} to {number}");
                    return true;
                }

                _logger.Debug($"{LogPrefix} Block not found: {blockName}");
                return false;
            }
            catch (Exception ex)
            {
                _logger.Error(ex, $"{LogPrefix} Error occurred while changing block number: {blockName}");
                return false;
            }
        }

        /// <inheritdoc />
        public bool CreateListFB(PlcSoftware plcSoftware, List<dataBlock> dataBlock)
        {
            if (plcSoftware == null) throw new ArgumentNullException(nameof(plcSoftware));
            if (dataBlock == null) throw new ArgumentNullException(nameof(dataBlock));

            try
            {
                foreach (var group in plcSoftware.BlockGroup.Groups)
                {
                    if (group.Name.Contains("DB") && group.Name.Contains("@"))
                    {
                        foreach (var eqGroup in group.Groups)
                        {
                            foreach (var block in eqGroup.Blocks.OfType<InstanceDB>())
                            {
                                var pos = block.Name.IndexOf("iDB");
                                if (pos != -1)
                                {
                                    var eq = block.Name.Substring(4).Split('|');
                                    dataBlock.Add(new dataBlock
                                    {
                                        name = block.Name,
                                        instanceOfName = block.InstanceOfName,
                                        number = block.Number,
                                        nameFC = dataBlock.FirstOrDefault(item => item.name == block.Name.Substring(4))?.nameFC,
                                        group = group.Name,
                                        typeEq = eq.Length > 1 ? eq[0] : "",
                                        nameEq = eq.Length > 1 ? eq[1] : "",
                                    });
                                    _logger.Information($"{LogPrefix} Added block to list: {block.Name}");
                                }
                            }
                        }
                    }
                }

                _logger.Information($"{LogPrefix} Created list of FB.");
                return true;
            }
            catch (Exception ex)
            {
                _logger.Error(ex, $"{LogPrefix} Error occurred while creating list of FB.");
                return false;
            }
        }
        #endregion
        #region Public Methods (Source Management)
        /// <inheritdoc />
        public string GenerateSourceBlock(PlcSoftware plcSoftware, string blockName, string path, GenerateOptions generateOption = GenerateOptions.None)
        {
            if (plcSoftware == null) throw new ArgumentNullException(nameof(plcSoftware));
            if (string.IsNullOrEmpty(blockName)) throw new ArgumentException("Block name cannot be null or empty.", nameof(blockName));
            if (string.IsNullOrEmpty(path)) throw new ArgumentException("Path cannot be null or empty.", nameof(path));

            try
            {
                var block = GetBlock(plcSoftware, blockName);

                if (block == null)
                {
                    _logger.Debug($"{LogPrefix} Block not found: {blockName}");
                    return null;
                }

                string blockType = string.Empty;
                switch (block.ProgrammingLanguage)
                {
                    case ProgrammingLanguage.STL:
                        blockType = ".awl";
                        break;
                    case ProgrammingLanguage.SCL:
                        blockType = ".scl";
                        break;
                    case ProgrammingLanguage.DB:
                        blockType = ".db";
                        break;
                    default:
                        _logger.Warning($"{LogPrefix} Unsupported programming language for block: {blockName}");
                        return null;
                }

                var filename = Path.GetFullPath(Path.Combine(path, block.Name + blockType));
                if (File.Exists(filename))
                {
                    File.Delete(filename);
                    _logger.Information($"{LogPrefix} Deleted existing source file: {filename}");
                }

                var fileInfo = new FileInfo(filename);
                plcSoftware.ExternalSourceGroup.GenerateSource(new List<PlcBlock> { block }, fileInfo, generateOption);
                _logger.Information($"{LogPrefix} Generated source for block: {blockName}");

                return filename;
            }
            catch (Exception ex)
            {
                _logger.Error(ex, $"{LogPrefix} Error occurred while generating source for block: {blockName}");
                return null;
            }
        }

        /// <inheritdoc />
        public string GenerateSourceUDT(PlcSoftware plcSoftware, string typeName, string path)
        {
            if (plcSoftware == null) throw new ArgumentNullException(nameof(plcSoftware));
            if (string.IsNullOrEmpty(typeName)) throw new ArgumentException("Type name cannot be null or empty.", nameof(typeName));
            if (string.IsNullOrEmpty(path)) throw new ArgumentException("Path cannot be null or empty.", nameof(path));

            try
            {
                var udt = plcSoftware.TypeGroup.Types.Find(typeName);
                if (udt == null)
                {
                    _logger.Warning($"{LogPrefix} UDT not found: {typeName}");
                    return null;
                }

                var filename = Path.GetFullPath(Path.Combine(path, udt.Name + ".udt"));
                if (File.Exists(filename))
                {
                    File.Delete(filename);
                    _logger.Information($"{LogPrefix} Deleted existing source file: {filename}");
                }

                var fileInfo = new FileInfo(filename);
                plcSoftware.ExternalSourceGroup.GenerateSource(new List<PlcType> { udt }, fileInfo, GenerateOptions.WithDependencies);
                _logger.Information($"{LogPrefix} Generated source for UDT: {typeName}");

                return filename;
            }
            catch (Exception ex)
            {
                _logger.Error(ex, $"{LogPrefix} Error occurred while generating source for UDT: {typeName}");
                return null;
            }
        }

        /// <inheritdoc />
        public bool ClearSource(PlcSoftware plcSoftware)
        {
            if (plcSoftware == null) throw new ArgumentNullException(nameof(plcSoftware));

            try
            {
                while (plcSoftware.ExternalSourceGroup.ExternalSources.Count > 0)
                {
                    var source = plcSoftware.ExternalSourceGroup.ExternalSources[0];
                    var sourceName = source.Name;
                    source.Delete();
                    _logger.Information($"{LogPrefix} Deleted external source: {sourceName}");
                }

                _logger.Information($"{LogPrefix} All external sources cleared.");
                return true;
            }
            catch (Exception ex)
            {
                _logger.Error(ex, $"{LogPrefix} Error occurred while clearing external sources.");
                return false;
            }
        }

        /// <inheritdoc />
        public bool ImportSource(PlcSoftware plcSoftware, string blockName, string path, string type = null)
        {
            if (plcSoftware == null) throw new ArgumentNullException(nameof(plcSoftware));
            if (string.IsNullOrEmpty(blockName)) throw new ArgumentException("Block name cannot be null or empty.", nameof(blockName));
            if (string.IsNullOrEmpty(path)) throw new ArgumentException("Path cannot be null or empty.", nameof(path));

            try
            {
                var existingSource = plcSoftware.ExternalSourceGroup.ExternalSources.Find(blockName);
                if (existingSource != null)
                {
                    existingSource.Delete();
                    _logger.Information($"{LogPrefix} Deleted existing external source: {blockName}");
                }

                var filename = Path.GetFullPath(Path.Combine(path, blockName + type));
                if (plcSoftware.ExternalSourceGroup.ExternalSources.CreateFromFile(blockName, filename) == null)
                {
                    _logger.Warning($"{LogPrefix} Wrong importing source: {blockName}");
                }
                _logger.Information($"{LogPrefix} Imported source: {blockName}");

                return true;
            }
            catch (Exception ex)
            {
                _logger.Error(ex, $"{LogPrefix} Error occurred while importing source: {blockName}");
                return false;
            }
        }
        #endregion
        #region Public Methods (Import/Export)
        /// <inheritdoc />
        public bool ExportBlock(PlcSoftware plcSoftware, string blockName, string path)
        {
            if (plcSoftware == null) throw new ArgumentNullException(nameof(plcSoftware));
            if (string.IsNullOrEmpty(blockName)) throw new ArgumentException("Block name cannot be null or empty.", nameof(blockName));
            if (string.IsNullOrEmpty(path)) throw new ArgumentException("Path cannot be null or empty.", nameof(path));

            try
            {
                var block = plcSoftware.BlockGroup.Blocks.Find(blockName);
                if (block == null)
                {
                    _logger.Debug($"{LogPrefix} Block not found: {blockName}");
                    return false;
                }

                var filename = Path.GetFullPath(Path.Combine(path, block.Name + ".xml"));
                if (File.Exists(filename))
                {
                    File.Delete(filename);
                    _logger.Information($"{LogPrefix} Deleted existing export file: {filename}");
                }

                block.Export(new FileInfo(filename), ExportOptions.WithDefaults);
                _logger.Information($"{LogPrefix} Exported block: {blockName}");

                return true;
            }
            catch (Exception ex)
            {
                _logger.Error(ex, $"{LogPrefix} Error occurred while exporting block: {blockName}");
                return false;
            }
        }

        /// <inheritdoc />
        public bool ImportBlock(PlcSoftware plcSoftware, string blockName, string path, string groupName = null)
        {
            if (plcSoftware == null) throw new ArgumentNullException(nameof(plcSoftware));
            if (string.IsNullOrEmpty(blockName)) throw new ArgumentException("Block name cannot be null or empty.", nameof(blockName));
            if (string.IsNullOrEmpty(path)) throw new ArgumentException("Path cannot be null or empty.", nameof(path));

            try
            {
                var filename = Path.GetFullPath(Path.Combine(path, blockName + ".xml"));
                if (groupName != null)
                {
                    var group = plcSoftware.BlockGroup.Groups.Find(groupName);
                    if (group == null)
                    {
                        group = plcSoftware.BlockGroup.Groups.Create(groupName);
                        _logger.Information($"{LogPrefix} Created group: {groupName}");
                    }

                    group.Blocks.Import(new FileInfo(filename), ImportOptions.Override, SWImportOptions.IgnoreMissingReferencedObjects);
                }
                else
                {
                    plcSoftware.BlockGroup.Blocks.Import(new FileInfo(filename), ImportOptions.Override, SWImportOptions.IgnoreMissingReferencedObjects);
                }

                _logger.Information($"{LogPrefix} Imported block: {blockName}");
                return true;
            }
            catch (Exception ex)
            {
                _logger.Error(ex, $"{LogPrefix} Error occurred while importing block: {blockName}");
                return false;
            }
        }
        #endregion
        #region Public Methods (Block Retrieval)
        /// <inheritdoc />
        public PlcBlock GetBlock(PlcSoftware plcSoftware, string fullPathName)
        {
            if (plcSoftware == null) throw new ArgumentNullException(nameof(plcSoftware));
            if (string.IsNullOrEmpty(fullPathName)) throw new ArgumentException("Block name cannot be null or empty.", nameof(fullPathName));

            try
            {
                PlcBlockSystemGroup systemGroup = plcSoftware.BlockGroup;
                PlcBlockUserGroupComposition groupComposition = systemGroup.Groups;
                PlcBlockUserGroup myCreatedGroup = null;

                string[] hierarchy = fullPathName.Split('.');
                foreach (string currentGroupName in hierarchy.Take(hierarchy.Length - 1))
                {
                    myCreatedGroup = groupComposition.Find(currentGroupName);
                    if (myCreatedGroup == null)
                    {
                        _logger.Warning($"{LogPrefix} Group not found in hierarchy: {currentGroupName}");
                        return null;
                    }
                    groupComposition = myCreatedGroup.Groups;
                }

                return myCreatedGroup.Blocks.Find(hierarchy.Last());
            }
            catch (Exception ex)
            {
                _logger.Error(ex, $"{LogPrefix} Error occurred while retrieving block group: {fullPathName}");
                return null;
            }
        }
        #endregion
    }
}