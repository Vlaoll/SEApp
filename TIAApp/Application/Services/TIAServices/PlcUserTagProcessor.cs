// Ignore Spelling: Eq Plc Conf

using Microsoft.Extensions.DependencyInjection;
using seConfSW.Domain.Models;
using Serilog;
using Siemens.Engineering.SW;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace seConfSW.Services
{
    /// <summary>
    /// Processor for handling PLC user tags and constants creation in TIA Portal environment.
    /// Provides functionality to create user constants, equipment constants, and import tags from files.
    /// </summary>
    public class PlcUserTagProcessor
    {
        #region Constants
        /// <summary>
        /// Prefix for log messages from this processor
        /// </summary>
        private const string LogPrefix = "[TIA/Tag]";
        #endregion
        #region Properties
        /// <summary>
        /// Event that fires when a message needs to be updated in the UI
        /// </summary>
        public event EventHandler<string> MessageUpdated;

        /// <summary>
        /// Configuration service for accessing application settings
        /// </summary>
        private readonly IConfigurationService _configuration;

        /// <summary>
        /// Logger instance for recording operation details
        /// </summary>
        private readonly ILogger _logger;

        /// <summary>
        /// Tag manager for handling tag-related operations
        /// </summary>
        private readonly ITagManager _tagManager;

        /// <summary>
        /// Path where tag source files will be stored temporarily
        /// </summary>
        private readonly string _sourceTagPath;
        #endregion
        #region Constructor
        /// <summary>
        /// Initializes a new instance of the PlcUserTagProcessor class.
        /// Resolves dependencies from the service provider.
        /// </summary>
        /// <exception cref="ArgumentNullException">Thrown when required services are not available</exception>
        public PlcUserTagProcessor()
        {
            _logger = App.ServiceProvider.GetService<ILogger>() ?? throw new ArgumentNullException(nameof(_logger));
            _configuration = App.ServiceProvider.GetService<IConfigurationService>() ?? throw new ArgumentNullException(nameof(_configuration));
            _tagManager = App.ServiceProvider.GetService<ITagManager>() ?? throw new ArgumentNullException(nameof(_tagManager));

            _sourceTagPath = _configuration.SourceTagPath;
        }
        #endregion
        #region Public Methods
        /// <summary>
        /// Creates common user constants for the specified PLC software based on the provided PLC data
        /// </summary>
        /// <param name="plcSoftware">The PLC software instance where constants will be created</param>
        /// <param name="dataPLC">PLC data containing constant definitions</param>
        /// <returns>True if operation succeeded, false otherwise</returns>
        public bool CreateCommonUserConstants(PlcSoftware plcSoftware, dataPLC dataPLC)
        {
            try
            {
                _logger.Information($"{LogPrefix} Starting creation of user constants for PLC: {dataPLC.namePLC}");
                _tagManager.CreateUserConstant(plcSoftware, dataPLC.userConstant);
                _logger.Information($"{LogPrefix} Successfully created user constants for PLC: {dataPLC.namePLC}");
                return true;
            }
            catch (Exception ex)
            {
                _logger.Error($"{LogPrefix} Failed to create user constants for PLC {dataPLC.namePLC}: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// Creates equipment-specific constants for the specified PLC software based on the provided PLC data
        /// </summary>
        /// <param name="plcSoftware">The PLC software instance where constants will be created</param>
        /// <param name="dataPLC">PLC data containing equipment and constant definitions</param>
        /// <returns>True if operation succeeded, false otherwise</returns>
        public bool CreateEqConstants(PlcSoftware plcSoftware, dataPLC dataPLC)
        {
            try
            {
                _logger.Information($"{LogPrefix} Starting creation of equipment constants for PLC: {dataPLC.namePLC}");

                foreach (var blockDB in dataPLC.instanceDB)
                {
                    var equipment = dataPLC.Equipment.FirstOrDefault(type => type.typeEq == blockDB.typeEq);
                    if (equipment?.dataConstant?.Count > 0)
                    {
                        _logger.Information($"{LogPrefix} Creating constants for block: {blockDB.nameEq}");
                        _tagManager.CreateUserConstant(plcSoftware, equipment.dataConstant, blockDB.excelData, blockDB.nameEq);
                    }
                }

                _logger.Information($"{LogPrefix} Successfully created equipment constants for PLC: {dataPLC.namePLC}");
                return true;
            }
            catch (Exception ex)
            {
                _logger.Error($"{LogPrefix} Failed to create equipment constants for PLC {dataPLC.namePLC}: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// Creates tags in the specified PLC software by importing them from XML files generated from the provided PLC data
        /// </summary>
        /// <param name="plcSoftware">The PLC software instance where tags will be created</param>
        /// <param name="dataPLC">PLC data containing tag definitions</param>
        /// <returns>True if operation succeeded, false otherwise</returns>
        public bool CreateTagsFromFile(PlcSoftware plcSoftware, dataPLC dataPLC)
        {
            try
            {
                _logger.Information($"{LogPrefix} Starting creation of tags for PLC: {dataPLC.namePLC}");

                var sources = new Dictionary<string, dataImportTag>();

                // Prepare the temporary directory for tag files
                if (Directory.Exists(_sourceTagPath)) Directory.Delete(_sourceTagPath, true);
                Directory.CreateDirectory(_sourceTagPath);

                // Process each instance in the PLC data
                foreach (var instance in dataPLC.instanceDB)
                {
                    var equipment = dataPLC.Equipment.FirstOrDefault(item => item.typeEq == instance.typeEq);
                    if (equipment == null) continue;

                    // Process each tag definition for the equipment
                    foreach (var tag in equipment.dataTag)
                    {
                        // Check if tag should be created based on variant matching
                        if ((tag.variant.Count == 0)
                            || tag.variant.Contains("")
                            || instance.variant.Intersect(tag.variant).Any())
                        {
                            if (string.IsNullOrEmpty(tag.adress)) continue;

                            var link = Common.ModifyString(tag.link, instance.excelData);
                            bool isExist = _tagManager.FindTag(plcSoftware, link, tag.table).isExist;

                            if (!isExist)
                            {
                                // Initialize XML source file if not exists
                                if (!sources.ContainsKey(tag.table))
                                {
                                    var filename = new FileInfo(Path.Combine(_sourceTagPath, $"{tag.table}.xml")).FullName;
                                    sources[tag.table] = new dataImportTag
                                    {
                                        fileName = filename,
                                        table = tag.table,
                                        ID = 0,
                                        name = tag.table,
                                        code = new StringBuilder()
                                            .AppendLine("<?xml version=\"1.0\" encoding=\"utf-8\"?>")
                                            .AppendLine("<Document>")
                                            .AppendLine("\t<Engineering version=\"V19\" />")
                                    };
                                }

                                // Build XML structure for the tag
                                var source = sources[tag.table];
                                var comments = Common.ModifyString(tag.comment, instance.excelData);
                                int ID = source.ID;

                                source.code
                                    .AppendLine("\t\t\t<SW.Tags.PlcTag ID=\"" + ID++ + "\" CompositionName=\"Tags\">")
                                    .AppendLine("\t\t\t\t<AttributeList>")
                                    .AppendLine("\t\t\t\t\t<DataTypeName>" + tag.type + "</DataTypeName>")
                                    .AppendLine("\t\t\t\t\t<LogicalAddress>%" + tag.adress + "</LogicalAddress>")
                                    .AppendLine("\t\t\t\t\t<Name>" + link + "</Name>")
                                    .AppendLine("\t\t\t\t</AttributeList>")
                                    .AppendLine("\t\t\t\t<ObjectList>")
                                    .AppendLine("\t\t\t\t\t<MultilingualText ID=\"" + ID++ + "\" CompositionName=\"Comment\">")
                                    .AppendLine("\t\t\t\t\t\t<ObjectList>")
                                    .AppendLine("\t\t\t\t\t\t\t<MultilingualTextItem ID=\"" + ID++ + "\" CompositionName=\"Items\">")
                                    .AppendLine("\t\t\t\t\t\t\t\t<AttributeList>")
                                    .AppendLine("\t\t\t\t\t\t\t\t\t<Culture>en-US</Culture>")
                                    .AppendLine("\t\t\t\t\t\t\t\t\t<Text>" + comments + "</Text>")
                                    .AppendLine("\t\t\t\t\t\t\t\t</AttributeList>")
                                    .AppendLine("\t\t\t\t\t\t\t</MultilingualTextItem>")
                                    .AppendLine("\t\t\t\t\t\t</ObjectList>")
                                    .AppendLine("\t\t\t\t\t</MultilingualText>")
                                    .AppendLine("\t\t\t\t</ObjectList>")
                                    .AppendLine("\t\t\t</SW.Tags.PlcTag>");

                                source.ID = ID;
                            }
                        }
                    }
                }

                // Write all XML files and import tags
                foreach (var source in sources.Values)
                {
                    source.code.AppendLine("</Document>");

                    using (var sw = new StreamWriter(source.fileName, false))
                    {
                        sw.WriteLine(source.code.ToString());
                    }

                    _logger.Information($"{LogPrefix} Importing tags for table: {source.table}");
                    _tagManager.ImportTags(plcSoftware, source.fileName, _configuration.SourceTagPath, source.table);
                }

                _logger.Information($"{LogPrefix} Successfully created tags for PLC: {dataPLC.namePLC}");
                return true;
            }
            catch (Exception ex)
            {
                _logger.Error($"{LogPrefix} Failed to create tags for PLC {dataPLC.namePLC}: {ex.Message}");
                return false;
            }
        }
        #endregion
    }
}