// Ignore Spelling: plc Conf Eq

using seConfSW.Domain.Models;
using Serilog;
using Siemens.Engineering.SW.Tags;
using Siemens.Engineering.SW;
using System.Collections.Generic;
using System.IO;
using System;
using System.Linq;
using Siemens.Engineering;

namespace seConfSW.Services
{
    /// <summary>
    /// Implements tag and user constant management functionality for TIA Portal
    /// </summary>
    public class TagManager : ITagManager
    {
        #region Constants
        private const string LogPrefix = "[TIA/TM]";
        private const string DefaultGroupName = "@Eq_TagTables";
        #endregion
        #region Properties
        private readonly ILogger _logger;
        #endregion
        #region Constructors

        /// <summary>
        /// Initializes new instance of TagManager
        /// </summary>
        /// <param name="logger">Logger instance</param>
        /// <exception cref="ArgumentNullException">Thrown when logger is null</exception>
        public TagManager(ILogger logger)
        {
            _logger = logger ?? throw new ArgumentNullException(nameof(logger));
        }
        #endregion
        #region Public Methods

        /// <inheritdoc/>
        public void CreateTag(PlcSoftware plcSoftware, dataPLC dataPLC, string groupName = DefaultGroupName)
        {
            if (plcSoftware == null) throw new ArgumentNullException(nameof(plcSoftware));
            if (dataPLC == null) throw new ArgumentNullException(nameof(dataPLC));
            if (string.IsNullOrWhiteSpace(groupName)) throw new ArgumentNullException(nameof(groupName));

            try
            {
                var systemGroup = plcSoftware.TagTableGroup;
                var groupComposition = systemGroup.Groups;
                var group = GetOrCreateGroup(groupComposition, groupName);

                // Group equipment by type for faster access
                var equipmentByType = dataPLC.Equipment.ToDictionary(eq => eq.typeEq);

                foreach (var instance in dataPLC.instanceDB)
                {
                    if (equipmentByType.TryGetValue(instance.typeEq, out var equipment))
                    {
                        foreach (var tag in equipment.dataTag)
                        {
                            // Check if tag applies to this variant
                            if (tag.variant.Count == 0 || instance.variant.Intersect(tag.variant).Any())
                            {
                                if (!string.IsNullOrWhiteSpace(tag.adress))
                                {
                                    var tagTable = GetOrCreateTagTable(group, tag.table);
                                    var modifiedTagLink = Common.ModifyString(tag.link, instance.excelData);

                                    if (tagTable.Tags.Find(modifiedTagLink) == null)
                                    {
                                        tagTable.Tags.Create(modifiedTagLink, tag.type, tag.adress);
                                        _logger.Information($"{LogPrefix} Created tag: {modifiedTagLink}");
                                    }
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                _logger.Error($"{LogPrefix} Error: {ex.Message}");
            }
        }

        /// <inheritdoc/>
        public void ImportTags(PlcSoftware plcSoftware, string filename, string sourceTagPath, string table, string groupName = DefaultGroupName)
        {
            if (plcSoftware == null) throw new ArgumentNullException(nameof(plcSoftware));
            if (string.IsNullOrWhiteSpace(filename)) throw new ArgumentNullException(nameof(filename));
            if (string.IsNullOrWhiteSpace(table)) throw new ArgumentNullException(nameof(table));
            if (string.IsNullOrWhiteSpace(sourceTagPath)) throw new ArgumentNullException(nameof(sourceTagPath));

            var fullPath = Path.GetFullPath(Path.Combine(sourceTagPath, filename));

            try
            {
                var fileInfo = new FileInfo(fullPath);
                if (!fileInfo.Exists)
                {
                    _logger.Error($"{LogPrefix} File not found: {fullPath}");
                    return;
                }

                var systemGroup = plcSoftware.TagTableGroup;
                var groupComposition = systemGroup.Groups;
                var group = GetOrCreateGroup(groupComposition, groupName);

                var tagTable = GetOrCreateTagTable(group, table);
                tagTable.Tags.Import(fileInfo, ImportOptions.None);
                _logger.Information($"{LogPrefix} Updated tags for table {table}");
            }
            catch (Exception ex)
            {
                _logger.Error($"{LogPrefix} Error: {ex.Message}");
            }
        }

        /// <inheritdoc/>
        public void ImportTagTable(PlcSoftware plcSoftware, string filename, string sourceTagPath, string groupName = DefaultGroupName)
        {
            if (plcSoftware == null) throw new ArgumentNullException(nameof(plcSoftware));
            if (string.IsNullOrWhiteSpace(filename)) throw new ArgumentNullException(nameof(filename));
            if (string.IsNullOrWhiteSpace(sourceTagPath)) throw new ArgumentNullException(nameof(sourceTagPath));

            var fullPath = Path.GetFullPath(Path.Combine(sourceTagPath, filename));

            try
            {
                var fileInfo = new FileInfo(fullPath);
                if (!fileInfo.Exists)
                {
                    _logger.Error($"{LogPrefix} File not found: {fullPath}");
                    return;
                }

                var systemGroup = plcSoftware.TagTableGroup;
                var groupComposition = systemGroup.Groups;
                var group = GetOrCreateGroup(groupComposition, groupName);

                var tagTableComposition = group.TagTables;
                tagTableComposition.Import(fileInfo, ImportOptions.Override);
                _logger.Information($"{LogPrefix} Created tags from file {filename}");
            }
            catch (Exception ex)
            {
                _logger.Error($"{LogPrefix} Error: {ex.Message}");
            }
        }

        /// <inheritdoc/>
        public (dataTag tag, bool isExist) FindTag(PlcSoftware plcSoftware, string tag, string table, string groupName = DefaultGroupName)
        {
            if (plcSoftware == null) throw new ArgumentNullException(nameof(plcSoftware));
            if (string.IsNullOrWhiteSpace(tag)) throw new ArgumentNullException(nameof(tag));
            if (string.IsNullOrWhiteSpace(table)) throw new ArgumentNullException(nameof(table));

            var tagData = new dataTag();

            try
            {
                var systemGroup = plcSoftware.TagTableGroup;
                var groupComposition = systemGroup.Groups;
                var group = groupComposition.Find(groupName);

                if (group == null)
                {
                    _logger.Debug($"{LogPrefix} Group not found: {groupName}");
                    return (tagData, false);
                }

                var tagTable = group.TagTables.Find(table);
                if (tagTable == null)
                {
                    _logger.Debug($"{LogPrefix} Table not found: {table}");
                    return (tagData, false);
                }

                var foundTag = tagTable.Tags.Find(tag);
                if (foundTag == null)
                {
                    _logger.Debug($"{LogPrefix} Tag not found: {tag}");
                    return (tagData, false);
                }

                tagData.link = foundTag.Name;
                tagData.adress = foundTag.LogicalAddress;
                tagData.type = foundTag.DataTypeName;
                tagData.comment = foundTag.Comment.ToString();
                return (tagData, true);
            }
            catch (Exception ex)
            {
                _logger.Error($"{LogPrefix} Error: {ex.Message}");
                return (tagData, false);
            }
        }

        /// <inheritdoc/>
        public void CreateUserConstant(PlcSoftware plcSoftware, List<userConstant> userConstant)
        {
            if (plcSoftware == null) throw new ArgumentNullException(nameof(plcSoftware));
            if (userConstant == null || userConstant.Count == 0) throw new ArgumentNullException(nameof(userConstant));

            try
            {
                var systemGroup = plcSoftware.TagTableGroup;
                var tagTableComposition = systemGroup.TagTables;
                var tagTable = tagTableComposition.FirstOrDefault(item => item.IsDefault);

                if (tagTable != null)
                {
                    foreach (var item in userConstant)
                    {
                        if (tagTable.UserConstants.Find(item.name) == null)
                        {
                            tagTable.UserConstants.Create(item.name, item.type, item.value);
                            _logger.Information($"{LogPrefix} Created user constant: {item.name}");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                _logger.Error($"{LogPrefix} Error: {ex.Message}");
            }
        }

        /// <inheritdoc/>
        public void CreateUserConstant(PlcSoftware plcSoftware, List<userConstant> userConstant, List<excelData> excelData, string EqName, string groupName = DefaultGroupName)
        {
            if (plcSoftware == null) throw new ArgumentNullException(nameof(plcSoftware));
            if (userConstant == null || userConstant.Count == 0) throw new ArgumentNullException(nameof(userConstant));
            if (excelData == null) throw new ArgumentNullException(nameof(excelData));
            if (string.IsNullOrWhiteSpace(EqName)) throw new ArgumentNullException(nameof(EqName));
            if (string.IsNullOrWhiteSpace(groupName)) throw new ArgumentNullException(nameof(groupName));

            try
            {
                var systemGroup = plcSoftware.TagTableGroup;
                var groupComposition = systemGroup.Groups;
                var group = GetOrCreateGroup(groupComposition, groupName);
                var equipmentConstants = new List<userConstant>();

                foreach (var item in userConstant)
                {
                    equipmentConstants.Add(new userConstant()
                    {
                        name = Common.ModifyString(item.name, excelData),
                        type = item.type,
                        value = Common.ModifyString(item.value, excelData),
                        table = item.table,
                    });
                }

                foreach (var item in equipmentConstants)
                {
                    var tagTable = systemGroup.TagTables.FirstOrDefault(gr => gr.IsDefault);
                    if (!string.IsNullOrWhiteSpace(item.table))
                    {
                        tagTable = GetOrCreateTagTable(group, item.table);
                    }

                    if (tagTable != null && tagTable.UserConstants.Find(item.name) == null)
                    {
                        tagTable.UserConstants.Create(item.name, item.type, item.value);
                    }
                }

                _logger.Information($"{LogPrefix} Created equipment constants for: {EqName}");
            }
            catch (Exception ex)
            {
                _logger.Error($"{LogPrefix} Error: {ex.Message}");
            }
        }
        #endregion
        #region Private Helper Methods

        /// <summary>
        /// Gets existing group or creates new one if not found
        /// </summary>
        /// <param name="groupComposition">Group composition to search in</param>
        /// <param name="groupName">Name of the group to find/create</param>
        /// <returns>Existing or newly created group</returns>
        private PlcTagTableUserGroup GetOrCreateGroup(PlcTagTableUserGroupComposition groupComposition, string groupName)
        {
            var group = groupComposition.Find(groupName);
            if (group == null)
            {
                group = groupComposition.Create(groupName);
                _logger.Information($"{LogPrefix} Created group for symbol tables: {groupName}");
            }
            return group;
        }

        /// <summary>
        /// Gets existing tag table or creates new one if not found
        /// </summary>
        /// <param name="group">Group to search in</param>
        /// <param name="tableName">Name of the table to find/create</param>
        /// <returns>Existing or newly created tag table</returns>
        private PlcTagTable GetOrCreateTagTable(PlcTagTableUserGroup group, string tableName)
        {
            var tagTable = group.TagTables.Find(tableName);
            if (tagTable == null)
            {
                tagTable = group.TagTables.Create(tableName);
                _logger.Information($"{LogPrefix} Created symbol table: {tableName}");
            }
            return tagTable;
        }
        #endregion
    }
}