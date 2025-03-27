// Ignore Spelling: Conf Plc Fc

using seConfSW.Domain.Models;
using Serilog;
using Siemens.Engineering.SW;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System;

namespace seConfSW.Services
{
    /// <summary>
    /// Processor for PLC function blocks that handles creation, modification, and management
    /// of function blocks (FCs) and instance data blocks (DBs) in TIA Portal projects.
    /// </summary>
    public class PlcFunctionBlockProcessor
    {
        #region Constants

        private const string LogPrefix = "[TIA/FC]";              // Prefix for log messages
        private const int NumberScopeDb = 20;                      // Number range scope for DB blocks
        private const int NumberStartDb = 1000;                    // Starting number for DB blocks
        private const string searchStringVariant = "isVariant - "; // Search string for variant regions

        #endregion
        #region Properties - Services

        private readonly IConfigurationService _configuration;     // Service for configuration management
        private readonly ILogger _logger;                          // Logging service
        private readonly IProjectManager _projectManager;          // Project management service
        private readonly IPlcBlockManager _plcBlockManager;       // PLC block management service
        private readonly ICompilerManager _compilerProcess;       // Compiler process service
        private readonly IPlcSourceManager _plcSourceManager;     // PLC source management service
        private readonly IHierarchyManager _hierarchyManager;     // Hierarchy management service

        #endregion
        #region Properties - Paths

        private readonly string _sourcePath;      // Default source path for PLC blocks
        private readonly string _templatePath;    // Template path for FC generation
        private readonly string _sourceDBPath;    // Source path for DB blocks

        #endregion
        #region Constructor

        /// <summary>
        /// Initializes a new instance of the PlcFunctionBlockProcessor class.
        /// </summary>
        /// <param name="logger">Logger service</param>
        /// <param name="configuration">Configuration service</param>
        /// <param name="projectManager">Project manager service</param>
        /// <param name="plcBlockManager">PLC block manager service</param>
        /// <param name="compilerProcess">Compiler process service</param>
        /// <param name="plcSourceManager">PLC source manager service</param>
        /// <param name="hierarchyManager">Hierarchy manager service</param>
        public PlcFunctionBlockProcessor(
            ILogger logger,
            IConfigurationService configuration,
            IProjectManager projectManager,
            IPlcBlockManager plcBlockManager,
            ICompilerManager compilerProcess,
            IPlcSourceManager plcSourceManager,
            IHierarchyManager hierarchyManager)
        {
            _configuration = configuration ?? throw new ArgumentNullException(nameof(configuration));
            _logger = logger ?? throw new ArgumentNullException(nameof(logger));
            _hierarchyManager = hierarchyManager ?? throw new ArgumentNullException(nameof(hierarchyManager));
            _plcBlockManager = plcBlockManager ?? throw new ArgumentNullException(nameof(plcBlockManager));
            _compilerProcess = compilerProcess ?? throw new ArgumentNullException(nameof(compilerProcess));
            _projectManager = projectManager ?? throw new ArgumentNullException(nameof(projectManager));
            _plcSourceManager = plcSourceManager ?? throw new ArgumentNullException(nameof(plcSourceManager));

            _sourcePath = _configuration.DefaultSourcePath;
            _templatePath = _configuration.TemplatePath;
            _sourceDBPath = _configuration.SourceDBPath;
        }

        #endregion
        #region Public Methods

        /// <summary>
        /// Creates template FCs for extended equipment types in the PLC.
        /// </summary>
        /// <param name="plcSoftware">PLC software instance</param>
        /// <param name="dataPLC">PLC configuration data</param>
        /// <returns>True if successful, false otherwise</returns>
        public bool CreateTemplateFcForExtendedType(PlcSoftware plcSoftware, dataPLC dataPLC)
        {
            var sources = new Dictionary<string, FileInfo>();
            var files = new Dictionary<string, FileInfo>();
            var templatePath = _configuration.TemplatePath;

            PrepareTemplateDirectory(templatePath);
            _logger.Information($"{LogPrefix} Starting creation of template FCs for PLC: {dataPLC.namePLC}");

            try
            {
                var extendedTypes = dataPLC.Equipment
                    .Where(t => t.isExtended)
                    .ToLookup(t => t.typeEq);

                var orderDataExcelListBlocks = dataPLC.instanceDB
                    .Where(e => extendedTypes.Contains(e.typeEq))
                    .OrderBy(o => o.number)
                    .ToList();
                var equipments = dataPLC.Equipment.Where(t => t.isExtended);
                GenerateTemplateSources(plcSoftware, equipments, templatePath);
                ProcessTemplateFCs(orderDataExcelListBlocks, files, sources, plcSoftware, dataPLC);
                CreateFCsFromFiles(plcSoftware, files);
                UpdateFCsFromSources(plcSoftware, sources, dataPLC);

                Cleanup(plcSoftware);
                _logger.Information($"{LogPrefix} Successfully created template FCs for PLC: {dataPLC.namePLC}");
                return true;
            }
            catch (Exception ex)
            {
                _logger.Error($"{LogPrefix} Failed to create template FCs for PLC {dataPLC.namePLC}: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// Edits FCs based on Excel data for all blocks in the PLC.
        /// </summary>
        /// <param name="plcSoftware">PLC software instance</param>
        /// <param name="dataPLC">PLC configuration data</param>
        /// <param name="listDataPLC">List of PLC configuration data</param>
        /// <param name="closeProject">Flag to close project after operation</param>
        /// <param name="saveProject">Flag to save project after operation</param>
        /// <param name="compileProject">Flag to compile project after operation</param>
        /// <returns>True if successful, false otherwise</returns>
        public bool EditFCFromExcelCallAllBlocks(PlcSoftware plcSoftware, dataPLC dataPLC, List<dataPLC> listDataPLC,
            bool closeProject = false, bool saveProject = false, bool compileProject = false)
        {
            var sources = new Dictionary<string, FileInfo>();
            _logger.Information($"{LogPrefix} Starting editing FCs for PLC: {dataPLC.namePLC}");

            try
            {
                var orderDataExcelListBlocks = GetNonExtendedBlocks(dataPLC);
                ProcessBlocks(orderDataExcelListBlocks, plcSoftware, dataPLC, listDataPLC, sources);
                UpdateFCs(plcSoftware, dataPLC, sources);
                FinalizeProject(plcSoftware, dataPLC, compileProject, saveProject, closeProject);

                _logger.Information($"{LogPrefix} Successfully edited FCs for PLC: {dataPLC.namePLC}");
                return true;
            }
            catch (Exception ex)
            {
                _logger.Error($"{LogPrefix} Failed to edit FCs for PLC {dataPLC.namePLC}: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// Creates instance blocks for the PLC based on configuration data.
        /// </summary>
        /// <param name="plcSoftware">PLC software instance</param>
        /// <param name="dataPLC">PLC configuration data</param>
        /// <returns>True if successful, false otherwise</returns>
        public bool CreateInstanceBlocks(PlcSoftware plcSoftware, dataPLC dataPLC)
        {
            _logger.Information($"{LogPrefix} Starting creation of instance blocks for PLC: {dataPLC.namePLC}");

            try
            {
                foreach (var blockDB in dataPLC.instanceDB)
                {
                    ProcessInstanceBlock(plcSoftware, dataPLC, blockDB);
                }

                _logger.Information($"{LogPrefix} Successfully created instance blocks for PLC: {dataPLC.namePLC}");
                return true;
            }
            catch (Exception ex)
            {
                _logger.Error($"{LogPrefix} Failed to create instance blocks for PLC {dataPLC.namePLC}: {ex.Message}");
                return false;
            }
        }

        #endregion
        #region Private Helper Methods - Template FC Processing

        /// <summary>
        /// Prepares the template directory by cleaning and recreating it.
        /// </summary>
        /// <param name="templatePath">Path to the template directory</param>
        private void PrepareTemplateDirectory(string templatePath)
        {
            if (Directory.Exists(templatePath)) Directory.Delete(templatePath, true);
            Directory.CreateDirectory(templatePath);
        }

        /// <summary>
        /// Generates template sources for extended equipment types.
        /// </summary>
        /// <param name="plcSoftware">PLC software instance</param>
        /// <param name="equipment">List of equipment data</param>
        /// <param name="templatePath">Path to store generated templates</param>
        private void GenerateTemplateSources(PlcSoftware plcSoftware, IEnumerable<dataEq> equipment, string templatePath)
        {
            foreach (var type in equipment)
            {
                var template = type.FB.FirstOrDefault();
                if (template == null) continue;

                _logger.Information($"{LogPrefix} Generating source for template: {template.name}");
                _plcSourceManager.GenerateSourceBlock(plcSoftware, $"{template.group}.{template.name}", templatePath);
            }
        }

        /// <summary>
        /// Processes template FCs for extended equipment types.
        /// </summary>
        /// <param name="dataBlocks">List of data blocks to process</param>
        /// <param name="files">Dictionary to store generated files</param>
        /// <param name="sources">Dictionary to store source files</param>
        /// <param name="plcSoftware">PLC software instance</param>
        /// <param name="dataPLC">PLC configuration data</param>
        private void ProcessTemplateFCs(List<dataBlock> dataBlocks, Dictionary<string, FileInfo> files,
           Dictionary<string, FileInfo> sources, PlcSoftware plcSoftware, dataPLC dataPLC)
        {
            foreach (var data in dataBlocks)
            {
                string templateFile = Path.Combine(_templatePath, $"{data.instanceOfName}.scl");
                var group = string.IsNullOrEmpty(data.nameFC)
                    ? data.group
                    : dataPLC?.dataFC?.FirstOrDefault(f => f.name == data.nameFC)?.group ?? data.group;

                var eqFileInfo = new FileInfo(Path.Combine(_templatePath, $"{group}.FC_Call-{data.typeEq}.scl"));
                string key = $"{group}.FC_Call-{data.typeEq}";
                if (!files.ContainsKey(key))
                {
                    files.Add(key, eqFileInfo);
                }

                _logger.Information($"{LogPrefix} Processing template FC for: {data.nameEq}");
                ProcessFile(templateFile, eqFileInfo.FullName, data);
                AddFunctionCall(sources, plcSoftware, dataPLC, data);
            }
        }

        /// <summary>
        /// Processes a template file to create a modified version for a specific block.
        /// </summary>
        /// <param name="inputFile">Input template file path</param>
        /// <param name="outputFile">Output file path</param>
        /// <param name="data">Block data for modification</param>
        private void ProcessFile(string inputFile, string outputFile, dataBlock data)
        {
            using (var sr = new StreamReader(inputFile))
            {
                using (var sw = new StreamWriter(outputFile, true))
                {
                    string line = sr.ReadLine();
                    if (line != null && line.Contains($"FUNCTION \"{data.instanceOfName}"))
                        sw.WriteLine($"FUNCTION \"{data.name}\" : Void");

                    while ((line = sr.ReadLine()) != null)
                    {
                        if (line.Contains("REGION") && !line.Contains("END_REGION"))
                        {
                            string lineNext = sr.ReadLine();
                            if (lineNext != null && lineNext.Contains("isVariant"))
                            {
                                ProcessVariantRegion(line, lineNext, data, sr, sw);
                            }
                            else
                            {
                                sw.WriteLine(Common.ModifyString(line, data.excelData, "@"));
                                sw.WriteLine(Common.ModifyString(lineNext, data.excelData, "@"));
                            }
                        }
                        else
                        {
                            sw.WriteLine(Common.ModifyString(line, data.excelData, "@"));
                        }
                    }
                }
            }
        }

        /// <summary>
        /// Adds a function call to the source dictionary for later processing.
        /// </summary>
        /// <param name="sources">Dictionary of source files</param>
        /// <param name="plcSoftware">PLC software instance</param>
        /// <param name="dataPLC">PLC configuration data</param>
        /// <param name="data">Block data for the function call</param>
        private void AddFunctionCall(Dictionary<string, FileInfo> sources, PlcSoftware plcSoftware, dataPLC dataPLC, dataBlock data)
        {
            var code = new[] {
                $"REGION Call template - \"{data.instanceOfName}\" for: {data.nameEq}",
                $"//{data.comment}",
                $" \"{data.name}\"();",
                "END_REGION",
                "END_FUNCTION"
            };

            ProcessSource(plcSoftware, dataPLC, data, sources, code);
        }

        /// <summary>
        /// Processes a variant region in a template file.
        /// </summary>
        /// <param name="line">Current line being processed</param>
        /// <param name="lineNext">Next line in the file</param>
        /// <param name="data">Block data containing variant information</param>
        /// <param name="sr">Stream reader for the input file</param>
        /// <param name="sw">Stream writer for the output file</param>
        private void ProcessVariantRegion(string line, string lineNext, dataBlock data, StreamReader sr, StreamWriter sw)
        {
            int n = lineNext.IndexOf(searchStringVariant) + searchStringVariant.Length;
            var variants = lineNext.Substring(n).Split(',').ToList();
            if (!data.variant.Intersect(variants).Any())
            {
                string l;
                while ((l = sr.ReadLine()) != null && l != "\tEND_REGION") { }
            }
            else
            {
                sw.WriteLine(Common.ModifyString(line, data.excelData, "@"));
                sw.WriteLine(lineNext);
            }
        }

        /// <summary>
        /// Creates FCs from generated template files.
        /// </summary>
        /// <param name="plcSoftware">PLC software instance</param>
        /// <param name="files">Dictionary of files to create FCs from</param>
        private void CreateFCsFromFiles(PlcSoftware plcSoftware, Dictionary<string, FileInfo> files)
        {
            foreach (var file in files)
            {
                _logger.Information($"{LogPrefix} Creating template FC: {file.Key}");
                var group = _hierarchyManager.CreateBlockGroup(plcSoftware, file.Key);
                if (group != null)
                {
                    var result = _plcBlockManager.CreateFC(plcSoftware, file.Key, 0, file.Value.DirectoryName, group, file.Value.Extension);
                }
            }
        }

        /// <summary>
        /// Updates FCs from generated source files.
        /// </summary>
        /// <param name="plcSoftware">PLC software instance</param>
        /// <param name="sources">Dictionary of source files</param>
        /// <param name="dataPLC">PLC configuration data</param>
        private void UpdateFCsFromSources(PlcSoftware plcSoftware, Dictionary<string, FileInfo> sources, dataPLC dataPLC)
        {
            foreach (var source in sources)
            {
                var fcData = dataPLC.dataFC.First(f => f.name == source.Key);
                _logger.Information($"{LogPrefix} Updating FC: {source.Key}");
                var group = _hierarchyManager.CreateBlockGroup(plcSoftware, fcData.group);
                if (group != null)
                {
                    var result = _plcBlockManager.CreateFC(plcSoftware, source.Key, fcData.number, source.Value.DirectoryName, group, source.Value.Extension);
                }
            }
        }

        /// <summary>
        /// Cleans up temporary resources after template processing.
        /// </summary>
        /// <param name="plcSoftware">PLC software instance</param>
        private void Cleanup(PlcSoftware plcSoftware)
        {
            _logger.Information($"{LogPrefix} Cleaning up temporary sources");
            _plcSourceManager.ClearSource(plcSoftware);
            _logger.Information($"{LogPrefix} Deleting temporary group @Template");
            plcSoftware.BlockGroup.Groups.Find("@Template").Delete();
        }

        #endregion
        #region Private Helper Methods - FC Editing

        /// <summary>
        /// Gets non-extended blocks from PLC configuration data.
        /// </summary>
        /// <param name="dataPLC">PLC configuration data</param>
        /// <returns>List of non-extended blocks</returns>
        private List<dataBlock> GetNonExtendedBlocks(dataPLC dataPLC)
        {
            return dataPLC.instanceDB
                .Where(element => dataPLC.Equipment.FirstOrDefault(type => type.typeEq == element.typeEq) != null &&
                                  dataPLC.Equipment.FirstOrDefault(type => type.typeEq == element.typeEq).isExtended == false)
                .OrderBy(order => order.number)
                .ToList();
        }

        /// <summary>
        /// Processes blocks for FC editing.
        /// </summary>
        /// <param name="blocks">List of blocks to process</param>
        /// <param name="plcSoftware">PLC software instance</param>
        /// <param name="dataPLC">PLC configuration data</param>
        /// <param name="listDataPLC">List of PLC configuration data</param>
        /// <param name="sources">Dictionary to store source files</param>
        private void ProcessBlocks(List<dataBlock> blocks, PlcSoftware plcSoftware,
           dataPLC dataPLC, List<dataPLC> listDataPLC, Dictionary<string, FileInfo> sources)
        {
            foreach (var data in blocks)
            {
                if (string.IsNullOrEmpty(data.typeEq) || string.IsNullOrEmpty(data.nameEq)) continue;

                var eqType = listDataPLC
                    .First(plc => plc.namePLC == dataPLC.namePLC)
                    .Equipment
                    .First(eq => eq.typeEq == data.typeEq);

                var io = BuildIOString(eqType.dataTag, data.excelData);
                var param = BuildParameterString(eqType.dataParameter, data);

                var code = new[] {
                $"REGION Call FB - \"{data.instanceOfName}\" for: {data.nameEq}",
                $"//Call functional block - {data.instanceOfName} for: {data.nameEq}",
                $"//{data.comment}",
                $"\"{data.name}\"{io.ToString()}",
                param.ToString(),
                "END_REGION",
                "END_FUNCTION"};

                ProcessSource(plcSoftware, dataPLC, data, sources, code);
            }
        }

        /// <summary>
        /// Builds the IO string for a block call.
        /// </summary>
        /// <param name="tags">List of tags for the block</param>
        /// <param name="excelData">Excel data for tag modification</param>
        /// <returns>StringBuilder with formatted IO string</returns>
        private StringBuilder BuildIOString(List<dataTag> tags, List<excelData> excelData)
        {
            var io = new StringBuilder("(");
            int tagCount = tags.Count;

            foreach (var tag in tags)
            {
                string line = $"{tag.name}:={Common.ModifyString(tag.link, excelData)}";
                io.Append(tagCount == 1 ? line : $"{line},");
                if (tagCount > 1) io.AppendLine();
                tagCount--;
            }
            io.Append(");");

            return io;
        }

        /// <summary>
        /// Builds the parameter string for a block call.
        /// </summary>
        /// <param name="parameters">List of parameters for the block</param>
        /// <param name="data">Block data</param>
        /// <returns>StringBuilder with formatted parameter string</returns>
        private StringBuilder BuildParameterString(List<dataParameter> parameters, dataBlock data)
        {
            var param = new StringBuilder();
            if (parameters.Count <= 1) return param;

            param.AppendLine($"//Parameters for block: {data.nameEq}");
            foreach (var parameter in parameters)
            {
                if (parameter.type == "I")
                    param.AppendLine($"\"{data.name}\".{parameter.name}:={Common.ModifyString(parameter.link, data.excelData)};");
                else if (parameter.type == "O")
                    param.AppendLine($"{Common.ModifyString(parameter.link, data.excelData)}:=\"{data.name}\".{parameter.name};");
            }
            return param;
        }

        /// <summary>
        /// Updates FCs from modified source files.
        /// </summary>
        /// <param name="plcSoftware">PLC software instance</param>
        /// <param name="dataPLC">PLC configuration data</param>
        /// <param name="sources">Dictionary of source files</param>
        private void UpdateFCs(PlcSoftware plcSoftware, dataPLC dataPLC, Dictionary<string, FileInfo> sources)
        {
            foreach (var source in sources)
            {
                var fcData = dataPLC.dataFC.First(fc => fc.name == source.Key);
                _logger.Information($"{LogPrefix} Updating FC: {source.Key}");
                var group = _hierarchyManager.CreateBlockGroup(plcSoftware, fcData.group);
                if (group != null)
                {
                    var result = _plcBlockManager.CreateFC(plcSoftware, source.Key, fcData.number, source.Value.DirectoryName, group, source.Value.Extension);
                }
            }
        }

        /// <summary>
        /// Finalizes the project after FC editing (compile, save, close).
        /// </summary>
        /// <param name="plcSoftware">PLC software instance</param>
        /// <param name="dataPLC">PLC configuration data</param>
        /// <param name="compileProject">Flag to compile project</param>
        /// <param name="saveProject">Flag to save project</param>
        /// <param name="closeProject">Flag to close project</param>
        private void FinalizeProject(PlcSoftware plcSoftware, dataPLC dataPLC,
           bool compileProject, bool saveProject, bool closeProject)
        {
            _logger.Information($"{LogPrefix} Clearing temporary sources");
            _plcSourceManager.ClearSource(plcSoftware);

            if (compileProject)
            {
                _logger.Information($"{LogPrefix} Compiling project for PLC: {dataPLC.namePLC}");
                _compilerProcess.Compile(dataPLC.namePLC, _projectManager.WorkProject);
            }
            if (saveProject)
            {
                _logger.Information($"{LogPrefix} Saving project");
                _projectManager.SaveProject();
            }
            if (closeProject)
            {
                _logger.Information($"{LogPrefix} Closing project");
                _projectManager.CloseProject();
            }
        }

        #endregion
        #region Private Helper Methods - Instance Block Processing

        /// <summary>
        /// Processes an instance block (creates either extended or simple instance DB).
        /// </summary>
        /// <param name="plcSoftware">PLC software instance</param>
        /// <param name="dataPLC">PLC configuration data</param>
        /// <param name="blockDB">Block data to process</param>
        private void ProcessInstanceBlock(PlcSoftware plcSoftware, dataPLC dataPLC, dataBlock blockDB)
        {
            int numberDB = CalculateDBNumber(blockDB.number);
            var equipment = GetEquipmentByType(dataPLC, blockDB.typeEq);

            if (equipment != null && equipment.isExtended)
            {
                CreateExtendedInstanceBlocks(plcSoftware, blockDB, equipment, numberDB);
            }
            else
            {
                CreateSimpleInstanceBlock(plcSoftware, blockDB, numberDB);
            }
        }

        /// <summary>
        /// Calculates the DB number based on the block number.
        /// </summary>
        /// <param name="blockNumber">Block number from configuration</param>
        /// <returns>Calculated DB number</returns>
        private int CalculateDBNumber(int blockNumber)
        {
            return blockNumber != 0 ? (blockNumber - 1) * NumberScopeDb + NumberStartDb : 0;
        }

        /// <summary>
        /// Gets equipment data by type from PLC configuration.
        /// </summary>
        /// <param name="dataPLC">PLC configuration data</param>
        /// <param name="typeEq">Equipment type to find</param>
        /// <returns>Equipment data or null if not found</returns>
        private dataEq GetEquipmentByType(dataPLC dataPLC, string typeEq)
        {
            return dataPLC.Equipment.FirstOrDefault(type => type.typeEq == typeEq);
        }

        /// <summary>
        /// Creates extended instance blocks for a DB.
        /// </summary>
        /// <param name="plcSoftware">PLC software instance</param>
        /// <param name="blockDB">Block data</param>
        /// <param name="equipment">Equipment data</param>
        /// <param name="baseNumber">Base number for DB numbering</param>
        private void CreateExtendedInstanceBlocks(PlcSoftware plcSoftware, dataBlock blockDB, dataEq equipment, int baseNumber)
        {
            foreach (var item in equipment.dataExtSupportBlock)
            {
                if (IsVariantMatch(item.variant, blockDB.variant))
                {
                    string name = Common.ModifyString(item.name, blockDB.excelData);
                    _logger.Information($"{LogPrefix} Creating extended instance DB: {name}");
                    var group = _hierarchyManager.CreateBlockGroup(plcSoftware, blockDB.group + '.' + blockDB.typeEq);
                    if (group != null)
                    {
                        _plcBlockManager.CreateInstanceDB(plcSoftware, name, baseNumber + item.number, item.instanceOfName, _sourceDBPath, group);
                    }
                }
            }
        }

        /// <summary>
        /// Checks if block variants match supported variants.
        /// </summary>
        /// <param name="supportVariants">List of supported variants</param>
        /// <param name="blockVariants">List of block variants</param>
        /// <returns>True if variants match, false otherwise</returns>
        private bool IsVariantMatch(List<string> supportVariants, List<string> blockVariants)
        {
            return supportVariants.Count == 0 || supportVariants.Contains("") || blockVariants.Intersect(supportVariants).Any();
        }

        /// <summary>
        /// Creates a simple instance block.
        /// </summary>
        /// <param name="plcSoftware">PLC software instance</param>
        /// <param name="blockDB">Block data</param>
        /// <param name="numberDB">Number for the DB</param>
        private void CreateSimpleInstanceBlock(PlcSoftware plcSoftware, dataBlock blockDB, int numberDB)
        {
            _logger.Information($"{LogPrefix} Creating instance DB: {blockDB.name}");
            var group = _hierarchyManager.CreateBlockGroup(plcSoftware, blockDB.group + '.' + blockDB.typeEq);
            if (group != null)
            {
                _plcBlockManager.CreateInstanceDB(plcSoftware, blockDB.name, numberDB, blockDB.instanceOfName, _sourceDBPath, group);
            }
        }

        #endregion
        #region Private Helper Methods - Source Processing

        /// <summary>
        /// Processes source files for FC generation/updating.
        /// </summary>
        /// <param name="plcSoftware">PLC software instance</param>
        /// <param name="dataPLC">PLC configuration data</param>
        /// <param name="data">Block data</param>
        /// <param name="sources">Dictionary of source files</param>
        /// <param name="code">Code to add to the source</param>
        private void ProcessSource(PlcSoftware plcSoftware, dataPLC dataPLC,
         dataBlock data, Dictionary<string, FileInfo> sources, string[] code)
        {
            FileInfo sourceFileInfo;
            string sourcePath = string.Empty;
            if (!sources.TryGetValue(data.nameFC, out sourceFileInfo))
            {
                var dateFC = dataPLC.dataFC.First(f => f.name == data.nameFC);
                var block = _plcBlockManager.GetBlock(plcSoftware, $"{dateFC.group}.{dateFC.name}");

                if (block == null)
                {
                    var group = _hierarchyManager.CreateBlockGroup(plcSoftware, dateFC.group);
                    if (group != null)
                    {
                        var result = _plcBlockManager.CreateFC(plcSoftware, dateFC.name, dateFC.number,
                        dateFC.code.AppendLine("END_FUNCTION").ToString(), _sourcePath, group, codeType: ".scl");
                    }
                }

                _logger.Information($"{LogPrefix} Generating source for FC: {dateFC.name}");
                sourcePath = _plcSourceManager.GenerateSourceBlock(plcSoftware, $"{dateFC.group}.{dateFC.name}", _sourcePath);
                if (!string.IsNullOrEmpty(sourcePath))
                {
                    sources.Add(data.nameFC, new FileInfo(sourcePath));
                }
            }
            else
            {
                sourcePath = sourceFileInfo?.FullName;
            }
            if (!string.IsNullOrEmpty(sourcePath))
            {
                var path = DeleteBlockRegion(sourcePath, data.nameEq);
                AddBlockRegion(path, code, data.nameEq);
            }
        }

        /// <summary>
        /// Deletes a block region from a source file.
        /// </summary>
        /// <param name="filename">Source file path</param>
        /// <param name="eqName">Equipment name to identify region</param>
        /// <returns>Path to the modified file</returns>
        private string DeleteBlockRegion(string filename, string eqName)
        {
            var tempFilename = Path.GetTempFileName();
            try
            {
                using (var sr = new StreamReader(filename))
                {
                    using (var sw = new StreamWriter(tempFilename))
                    {
                        string line;
                        bool skipRegion = false;

                        while ((line = sr.ReadLine()) != null)
                        {
                            if (line.Contains("REGION Call") && line.Contains(eqName))
                            {
                                skipRegion = true;
                                _logger.Information($"{LogPrefix} Deleted region for block: {eqName}");
                            }

                            if (!skipRegion)
                            {
                                sw.WriteLine(line);
                            }

                            if (skipRegion && line.Contains("END_REGION"))
                            {
                                skipRegion = false;
                            }
                        }
                    }
                }

                File.Delete(filename);
                File.Move(tempFilename, filename);
                return filename;
            }
            catch (Exception e)
            {
                _logger.Error($"{LogPrefix} Failed to delete region for block {eqName}: {e.Message}");
                return filename;
            }
            finally
            {
                if (File.Exists(tempFilename)) File.Delete(tempFilename);
            }
        }

        /// <summary>
        /// Adds a block region to a source file.
        /// </summary>
        /// <param name="filename">Source file path</param>
        /// <param name="newBlock">Code lines to add</param>
        /// <param name="eqName">Equipment name for logging</param>
        /// <returns>True if successful, false otherwise</returns>
        private bool AddBlockRegion(string filename, string[] newBlock, string eqName)
        {
            var tempFilename = Path.GetTempFileName();
            try
            {
                using (var sr = new StreamReader(filename))
                {
                    using (var sw = new StreamWriter(tempFilename))
                    {
                        string line;

                        while ((line = sr.ReadLine()) != null)
                        {
                            if (!line.Contains("END_FUNCTION"))
                            {
                                sw.WriteLine(line);
                            }
                            else
                            {
                                foreach (var item in newBlock)
                                {
                                    sw.WriteLine(item);
                                }
                                _logger.Information($"{LogPrefix} Added region for block: {eqName}");
                            }
                        }
                    }
                }

                File.Delete(filename);
                File.Move(tempFilename, filename);
                return true;
            }
            catch (Exception e)
            {
                _logger.Error($"{LogPrefix} Failed to add region for block {eqName}: {e.Message}");
                return false;
            }
            finally
            {
                if (File.Exists(tempFilename)) File.Delete(tempFilename);
            }
        }

        #endregion
    }
}