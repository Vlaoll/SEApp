// Ignore Spelling: plc Conf Eq Prj

using System;
using System.Linq;
using Serilog;
using Siemens.Engineering.SW;
using Siemens.Engineering.SW.Blocks;
using Siemens.Engineering.SW.Tags;
using Siemens.Engineering.SW.Types;

namespace seConfSW.Services
{
    /// <summary>
    /// Manages the creation and management of hierarchy elements (types, blocks, and tag tables) 
    /// in Siemens TIA Portal projects.
    /// Provides methods to create type groups, block groups, and tag tables in a hierarchical structure.
    /// </summary>
    public class HierarchyManager : IHierarchyManager
    {
        #region Constants

        /// <summary>
        /// Prefix for log messages from this class
        /// </summary>
        private const string LogPrefix = "[TIA/HM]";

        #endregion
        #region Properties

        /// <summary>
        /// Logger instance for recording information and errors
        /// </summary>
        private readonly ILogger _logger;

        #endregion
        #region Constructor

        /// <summary>
        /// Initializes a new instance of the <see cref="HierarchyManager"/> class.
        /// </summary>
        /// <param name="logger">The logger instance for recording operations</param>
        /// <exception cref="ArgumentNullException">Thrown when logger is null</exception>
        public HierarchyManager(ILogger logger)
        {
            _logger = logger ?? throw new ArgumentNullException(nameof(logger));
        }

        #endregion
        #region Public Methods

        /// <inheritdoc />
        public PlcTypeUserGroup CreateTypeGroup(PlcSoftware plcSoftware, string groupName)
        {
            if (plcSoftware == null) throw new ArgumentNullException(nameof(plcSoftware));
            if (string.IsNullOrEmpty(groupName)) throw new ArgumentException("Group name cannot be null or empty", nameof(groupName));

            try
            {
                PlcTypeSystemGroup systemGroup = plcSoftware.TypeGroup;
                PlcTypeUserGroupComposition groupComposition = systemGroup.Groups;
                PlcTypeUserGroup myCreatedGroup = null;

                // Process each level in the dot-separated hierarchy path
                foreach (string currentGroupName in groupName.Split('.'))
                {
                    myCreatedGroup = groupComposition.Find(currentGroupName);
                    if (myCreatedGroup == null)
                    {
                        myCreatedGroup = groupComposition.Create(currentGroupName);
                        _logger.Information($"{LogPrefix} Create group for types: {currentGroupName}");
                    }

                    groupComposition = myCreatedGroup.Groups; // Move to the next level in the hierarchy
                }

                return myCreatedGroup;
            }
            catch (Exception ex)
            {
                _logger.Error($"{LogPrefix} Error creating type group: {ex.Message}");
                return null;
            }
        }

        /// <inheritdoc />
        public PlcBlockUserGroup CreateBlockGroup(PlcSoftware plcSoftware, string groupName)
        {
            if (plcSoftware == null) throw new ArgumentNullException(nameof(plcSoftware));
            if (string.IsNullOrEmpty(groupName)) throw new ArgumentException("Group name cannot be null or empty", nameof(groupName));

            try
            {
                PlcBlockSystemGroup systemGroup = plcSoftware.BlockGroup;
                PlcBlockUserGroupComposition groupComposition = systemGroup.Groups;
                PlcBlockUserGroup myCreatedGroup = null;

                // Process each level in the dot-separated hierarchy path
                foreach (string currentGroupName in groupName.Split('.'))
                {
                    myCreatedGroup = groupComposition.Find(currentGroupName);

                    if (myCreatedGroup == null)
                    {
                        myCreatedGroup = groupComposition.Create(currentGroupName);
                        _logger.Information($"{LogPrefix} Create group for blocks: {currentGroupName}");
                    }
                    groupComposition = myCreatedGroup.Groups; // Move to the next level in the hierarchy 
                }

                return myCreatedGroup;
            }
            catch (Exception ex)
            {
                _logger.Error($"{LogPrefix} Error creating block group: {ex.Message}");
                return null;
            }
        }

        /// <inheritdoc />
        public PlcTagTable CreateTagTable(PlcSoftware plcSoftware, string groupName)
        {
            if (plcSoftware == null) throw new ArgumentNullException(nameof(plcSoftware));
            if (string.IsNullOrEmpty(groupName)) throw new ArgumentException("Group name cannot be null or empty", nameof(groupName));

            try
            {
                PlcTagTableSystemGroup systemGroup = plcSoftware.TagTableGroup;
                PlcTagTableUserGroupComposition groupComposition = systemGroup.Groups;
                PlcTagTableUserGroup myCreatedGroup = null;

                string[] hierarchy = groupName.Split('.');
                // Process all levels except the last one as group hierarchy
                foreach (string currentGroupName in hierarchy.Take(hierarchy.Length - 1))
                {
                    myCreatedGroup = groupComposition.Find(currentGroupName);

                    if (myCreatedGroup == null)
                    {
                        myCreatedGroup = groupComposition.Create(currentGroupName);
                        _logger.Information($"{LogPrefix} Create group for TagTable: {currentGroupName}");
                    }
                    groupComposition = myCreatedGroup.Groups; // Move to the next level in the hierarchy
                }

                // Create the tag table in the last group
                string tagTableName = hierarchy.Last();
                PlcTagTableComposition tagTableComposition = myCreatedGroup.TagTables;
                PlcTagTable myTagTable = tagTableComposition.Find(tagTableName) ?? tagTableComposition.Create(tagTableName);
                _logger.Information($"{LogPrefix} Created TagTable: {tagTableName}");

                return myTagTable;
            }
            catch (Exception ex)
            {
                _logger.Error($"{LogPrefix} Error creating tag table: {ex.Message}");
                return null;
            }
        }

        #endregion
    }
}