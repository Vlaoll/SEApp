// Ignore Spelling: Conf

using Siemens.Engineering.SW;
using Siemens.Engineering.SW.Blocks;
using Siemens.Engineering.SW.Tags;
using Siemens.Engineering.SW.Types;

namespace seConfSW.Services
{
    /// <summary>
    /// Defines the interface for managing hierarchy elements in TIA Portal projects.
    /// Provides methods for creating type groups, block groups, and tag tables.
    /// </summary>
    public interface IHierarchyManager
    {
        /// <summary>
        /// Creates a hierarchical structure of type groups in the specified PLC software.
        /// </summary>
        /// <param name="plcSoftware">The PLC software instance where groups will be created</param>
        /// <param name="groupName">Dot-separated path of group names (e.g., "Parent.Child.Grandchild")</param>
        /// <returns>The created or existing type group at the lowest level</returns>
        PlcTypeUserGroup CreateTypeGroup(PlcSoftware plcSoftware, string groupName);

        /// <summary>
        /// Creates a hierarchical structure of block groups in the specified PLC software.
        /// </summary>
        /// <param name="plcSoftware">The PLC software instance where groups will be created</param>
        /// <param name="groupName">Dot-separated path of group names (e.g., "Parent.Child.Grandchild")</param>
        /// <returns>The created or existing block group at the lowest level</returns>
        PlcBlockUserGroup CreateBlockGroup(PlcSoftware plcSoftware, string groupName);

        /// <summary>
        /// Creates a hierarchical structure of tag table groups and a tag table in the specified PLC software.
        /// The last part of the groupName is used as the tag table name.
        /// </summary>
        /// <param name="plcSoftware">The PLC software instance where groups and tag table will be created</param>
        /// <param name="groupName">Dot-separated path where the last part is the tag table name (e.g., "Parent.Child.TableName")</param>
        /// <returns>The created or existing tag table</returns>
        PlcTagTable CreateTagTable(PlcSoftware plcSoftware, string groupName);
    }
}