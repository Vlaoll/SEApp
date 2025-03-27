// Ignore Spelling: Plc Conf

using seConfSW.Domain.Models;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using Serilog;
using Siemens.Engineering.SW;
using Microsoft.Extensions.DependencyInjection;

namespace seConfSW.Services
{
    /// <summary>
    /// Processor for handling PLC data block operations including creation, updating, and importing.
    /// </summary>
    public class PlcDataBlockProcessor
    {
        #region Constants
        /// <summary>
        /// Prefix for log messages from this processor
        /// </summary>
        private const string LogPrefix = "[TIA/SuportDB]";
        #endregion
        #region Properties
        /// <summary>
        /// Logger instance for recording operations and errors
        /// </summary>
        private readonly ILogger _logger;

        /// <summary>
        /// Manager for handling hierarchy operations in PLC software
        /// </summary>
        private readonly IHierarchyManager _hierarchyManager;

        /// <summary>
        /// Manager for handling PLC block operations
        /// </summary>
        private readonly IPlcBlockManager _plcBlockManager;

        /// <summary>
        /// Service for accessing configuration settings
        /// </summary>
        private readonly IConfigurationService _configuration;

        /// <summary>
        /// Path where source files are stored
        /// </summary>
        private readonly string _sourcePath;
        #endregion
        #region Constructor
        /// <summary>
        /// Initializes a new instance of the PlcDataBlockProcessor
        /// </summary>
        /// <param name="sourcePath">Path to the directory where source files will be stored</param>
        /// <exception cref="ArgumentNullException">Thrown when any required service or parameter is null</exception>
        public PlcDataBlockProcessor(string sourcePath)
        {
            _logger = App.ServiceProvider.GetService<ILogger>() ?? throw new ArgumentNullException(nameof(_logger));
            _hierarchyManager = App.ServiceProvider.GetService<IHierarchyManager>() ?? throw new ArgumentNullException(nameof(_hierarchyManager));
            _plcBlockManager = App.ServiceProvider.GetService<IPlcBlockManager>() ?? throw new ArgumentNullException(nameof(_plcBlockManager));
            _configuration = App.ServiceProvider.GetService<IConfigurationService>() ?? throw new ArgumentNullException(nameof(_configuration));
            _sourcePath = sourcePath ?? throw new ArgumentNullException(nameof(sourcePath));
        }
        #endregion
        #region Public Methods
        /// <summary>
        /// Main method to add values to PLC data blocks
        /// </summary>
        /// <param name="plcSoftware">PLC software instance to work with</param>
        /// <param name="dataPLC">PLC data containing blocks and values to process</param>
        /// <returns>True if operation completed successfully, false otherwise</returns>
        public bool AddValueToDataBlock(PlcSoftware plcSoftware, dataPLC dataPLC)
        {
            try
            {
                _logger.Information($"{LogPrefix} Starting to add values to data blocks for PLC: {dataPLC.namePLC}");

                var (dataDBValues, dataSupportBD) = CollectData(dataPLC);
                return ProcessDataBlocks(plcSoftware, dataSupportBD, dataDBValues);
            }
            catch (Exception ex)
            {
                _logger.Error($"{LogPrefix} Failed to add values to data block: {ex.Message}");
                return false;
            }
        }
        #endregion
        #region Private Helper Methods
        /// <summary>
        /// Collects and organizes data from PLC instances into dictionaries and lists
        /// </summary>
        /// <param name="dataPLC">PLC data containing instances and equipment</param>
        /// <returns>Tuple containing:
        /// - Dictionary of DB values (key: DB name, value: dictionary of name-type pairs)
        /// - List of support block data</returns>
        private (Dictionary<string, Dictionary<string, string>>, List<dataSupportBD>) CollectData(dataPLC dataPLC)
        {
            var dataDBValues = new Dictionary<string, Dictionary<string, string>>();
            var dataSupportBD = new List<dataSupportBD>();

            foreach (var instance in dataPLC.instanceDB)
            {
                var equipment = dataPLC.Equipment.FirstOrDefault(type => type.typeEq == instance.typeEq);
                if (equipment == null)
                {
                    _logger.Warning($"{LogPrefix} No equipment found for type: {instance.typeEq}");
                    continue;
                }

                // Process data block values
                foreach (var value in equipment.dataDataBlockValue)
                {
                    string name = Common.ModifyString(value.name, instance.excelData);
                    if (!dataDBValues.ContainsKey(value.DB))
                    {
                        dataDBValues.Add(value.DB, new Dictionary<string, string>());
                    }
                    dataDBValues[value.DB][name] = $" : {value.type};";
                }

                // Process support blocks
                foreach (var support in equipment.dataSupportBD)
                {
                    if (!dataSupportBD.Any(item => item.name == support.name))
                    {
                        dataSupportBD.Add(new dataSupportBD
                        {
                            name = support.name,
                            number = support.number,
                            group = support.group,
                            path = support.path,
                            isType = support.isType,
                            isRetain = support.isRetain,
                            isOptimazed = support.isOptimazed
                        });
                    }
                }
            }

            return (dataDBValues, dataSupportBD);
        }

        /// <summary>
        /// Processes all data blocks by creating groups, generating sources, and updating values
        /// </summary>
        /// <param name="plcSoftware">PLC software instance</param>
        /// <param name="dataSupportBD">List of support blocks to process</param>
        /// <param name="dataDBValues">Dictionary containing data block values</param>
        /// <returns>True if all blocks were processed successfully</returns>
        private bool ProcessDataBlocks(PlcSoftware plcSoftware, List<dataSupportBD> dataSupportBD, Dictionary<string, Dictionary<string, string>> dataDBValues)
        {
            foreach (var db in dataSupportBD)
            {
                _logger.Information($"{LogPrefix} Creating block group for {db.group}");
                var group = _hierarchyManager.CreateBlockGroup(plcSoftware, db.group);

                _logger.Information($"{LogPrefix} Generating source block for {db.group}.{db.name}");
                string filename = _plcBlockManager.GenerateSourceBlock(plcSoftware, $"{db.group}.{db.name}", _configuration.DefaultSourcePath);

                if (string.IsNullOrEmpty(filename))
                {
                    filename = CreateNewDataBlock(db);
                }
                if (!string.IsNullOrEmpty(filename))
                {
                    UpdateExistingDataBlock(db, filename, dataDBValues);
                    ImportAndUpdateBlock(plcSoftware, db, filename);
                }
            }

            return true;
        }

        /// <summary>
        /// Creates a new empty data block file with basic structure
        /// </summary>
        /// <param name="db">Support block data containing configuration</param>
        /// <returns>Path to the created file</returns>
        private string CreateNewDataBlock(dataSupportBD db)
        {
            var codeSB = new StringBuilder();
            codeSB.AppendLine($"DATA_BLOCK \"{db.name}\"")
                  .AppendLine(db.isOptimazed ? "{ S7_Optimized_Access := 'TRUE' }" : "{ S7_Optimized_Access := 'FALSE' }")
                  .AppendLine("VERSION : 0.1")
                  .AppendLine(db.isRetain ? "" : "NON_RETAIN")
                  .AppendLine("STRUCT")
                  .AppendLine("END_STRUCT;")
                  .AppendLine("BEGIN")
                  .AppendLine("END_DATA_BLOCK");

            string filename = $"{_sourcePath}{db.name}.db";
            using (var sw = new StreamWriter(filename))
            {
                sw.Write(codeSB.ToString().Replace("\r", ""));
            }

            _logger.Information($"{LogPrefix} Created empty source for data block: {db.name}");
            return filename;
        }

        /// <summary>
        /// Updates an existing data block file with new values
        /// </summary>
        /// <param name="db">Support block data containing configuration</param>
        /// <param name="filename">Path to the existing file</param>
        /// <param name="dataDBValues">Dictionary containing values to update</param>
        private void UpdateExistingDataBlock(dataSupportBD db, string filename, Dictionary<string, Dictionary<string, string>> dataDBValues)
        {
            var tempFilename = Path.GetTempFileName();
            using (var sr = new StreamReader(filename))
            using (var sw = new StreamWriter(tempFilename))
            {
                string line;
                while ((line = sr.ReadLine()) != null)
                {
                    if (line.Contains("STRUCT") || line.Contains("VAR RETAIN"))
                    {
                        ParseAndUpdateValues(sr, dataDBValues, db.name);
                        break;
                    }
                    else if (line.Contains("BEGIN"))
                    {
                        break;
                    }
                    sw.WriteLine(line);
                }

                WriteUpdatedStructure(sw, db, dataDBValues);
            }

            File.Delete(filename);
            File.Move(tempFilename, filename);
            _logger.Information($"{LogPrefix} Generated new source for data block: {db.name}");
        }

        /// <summary>
        /// Parses existing values from a data block and updates the values dictionary
        /// </summary>
        /// <param name="sr">StreamReader for the source file</param>
        /// <param name="dataDBValues">Dictionary to update with parsed values</param>
        /// <param name="dbName">Name of the data block being processed</param>
        private void ParseAndUpdateValues(StreamReader sr, Dictionary<string, Dictionary<string, string>> dataDBValues, string dbName)
        {
            string line;
            while ((line = sr.ReadLine()) != null && (!line.Contains("END_VAR") && !line.Contains("END_STRUCT;")))
            {
                var parts = line.Replace(" ", "").Split(':');
                if (parts.Length == 2)
                {
                    if (!dataDBValues.ContainsKey(dbName))
                    {
                        dataDBValues.Add(dbName, new Dictionary<string, string>());
                    }
                    dataDBValues[dbName][parts[0]] = $" : {parts[1]}";
                }
            }
        }

        /// <summary>
        /// Writes the updated structure to the data block file
        /// </summary>
        /// <param name="sw">StreamWriter for the target file</param>
        /// <param name="db">Support block data containing configuration</param>
        /// <param name="dataDBValues">Dictionary containing values to write</param>
        private void WriteUpdatedStructure(StreamWriter sw, dataSupportBD db, Dictionary<string, Dictionary<string, string>> dataDBValues)
        {
            if (db.isRetain)
            {
                sw.WriteLine("VAR RETAIN");
                if (dataDBValues.ContainsKey(db.name))
                {
                    foreach (var item in dataDBValues[db.name])
                    {
                        sw.WriteLine($"{item.Key}{item.Value}");
                    }
                }
                sw.WriteLine("END_VAR");
            }
            else
            {
                sw.WriteLine("STRUCT");
                if (dataDBValues.ContainsKey(db.name))
                {
                    foreach (var item in dataDBValues[db.name])
                    {
                        sw.WriteLine($"{item.Key}{item.Value}");
                    }
                }
                sw.WriteLine("END_STRUCT;");
            }
            sw.WriteLine("BEGIN");
            sw.WriteLine("END_DATA_BLOCK");
        }

        /// <summary>
        /// Imports and updates a block in the PLC software
        /// </summary>
        /// <param name="plcSoftware">PLC software instance</param>
        /// <param name="db">Support block data containing configuration</param>
        /// <param name="filename">Path to the source file</param>
        private void ImportAndUpdateBlock(PlcSoftware plcSoftware, dataSupportBD db, string filename)
        {
            _logger.Information($"{LogPrefix} Importing source for data block: {db.name}");
            _plcBlockManager.ClearSource(plcSoftware);
            var rImport = _plcBlockManager.ImportSource(plcSoftware, db.name, _sourcePath, ".db");

            if (rImport)
            {
                var group = _hierarchyManager.CreateBlockGroup(plcSoftware, db.group);
                if (group != null)
                {
                    var rGenerate = _plcBlockManager.GenerateBlock(plcSoftware, group);
                    if (rGenerate)
                    {
                        var rChange = _plcBlockManager.ChangeBlockNumber(plcSoftware, db.name, db.number, group);
                        if (rChange)
                        {
                            _logger.Information($"{LogPrefix} Importing source for data block: {db.name} - successful");
                            return;
                        }
                    }
                }
            }
            _logger.Warning($"{LogPrefix} Importing source for data block: {db.name} - wrong");
        }
        #endregion
    }
}