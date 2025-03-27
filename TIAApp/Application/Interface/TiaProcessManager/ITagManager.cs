// Ignore Spelling: plc Conf Eq

using seConfSW.Domain.Models;
using Siemens.Engineering.SW;
using System.Collections.Generic;

namespace seConfSW.Services
{
    /// <summary>
    /// Provides functionality for managing PLC tags and user constants in TIA Portal
    /// </summary>
    public interface ITagManager
    {
        /// <summary>
        /// Creates tags in PLC software based on configuration data
        /// </summary>
        /// <param name="plcSoftware">PLC software instance</param>
        /// <param name="dataPLC">Configuration data for PLC</param>
        /// <param name="groupName">Name of the tag group (default: "@Eq_TagTables")</param>
        void CreateTag(PlcSoftware plcSoftware, dataPLC dataPLC, string groupName = "@Eq_TagTables");

        /// <summary>
        /// Imports tags from file into specified tag table
        /// </summary>
        /// <param name="plcSoftware">PLC software instance</param>
        /// <param name="filename">Name of the import file</param>
        /// <param name="sourceTagPath">Path to the tag file</param>
        /// <param name="table">Target tag table name</param>
        /// <param name="groupName">Name of the tag group (default: "@Eq_TagTables")</param>
        void ImportTags(PlcSoftware plcSoftware, string filename, string sourceTagPath, string table, string groupName = "@Eq_TagTables");

        /// <summary>
        /// Imports entire tag table from file
        /// </summary>
        /// <param name="plcSoftware">PLC software instance</param>
        /// <param name="filename">Name of the import file</param>
        /// <param name="sourceTagPath">Path to the tag file</param>
        /// <param name="groupName">Name of the tag group (default: "@Eq_TagTables")</param>
        void ImportTagTable(PlcSoftware plcSoftware, string filename, string sourceTagPath, string groupName = "@Eq_TagTables");

        /// <summary>
        /// Finds tag in specified table and group
        /// </summary>
        /// <param name="plcSoftware">PLC software instance</param>
        /// <param name="tag">Tag name to find</param>
        /// <param name="table">Tag table name</param>
        /// <param name="groupName">Name of the tag group (default: "@Eq_TagTables")</param>
        /// <returns>Tuple containing tag data and existence flag</returns>
        (dataTag tag, bool isExist) FindTag(PlcSoftware plcSoftware, string tag, string table, string groupName = "@Eq_TagTables");

        /// <summary>
        /// Creates user constants in default tag table
        /// </summary>
        /// <param name="plcSoftware">PLC software instance</param>
        /// <param name="userConstant">List of user constants to create</param>
        void CreateUserConstant(PlcSoftware plcSoftware, List<userConstant> userConstant);

        /// <summary>
        /// Creates user constants with equipment-specific names
        /// </summary>
        /// <param name="plcSoftware">PLC software instance</param>
        /// <param name="userConstant">List of user constant templates</param>
        /// <param name="excelData">Equipment data for name substitution</param>
        /// <param name="EqName">Equipment name</param>
        /// <param name="groupName">Name of the tag group (default: "@Eq_TagTables")</param>
        void CreateUserConstant(PlcSoftware plcSoftware, List<userConstant> userConstant, List<excelData> excelData, string EqName, string groupName = "@Eq_TagTables");
    }
}