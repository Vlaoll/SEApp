// Ignore Spelling: Conf Prj plc

using Siemens.Engineering;
using Siemens.Engineering.Library;
using Siemens.Engineering.SW;

namespace seConfSW.Services
{
    /// <summary>
    /// Provides functionality for managing TIA Portal libraries including opening, updating, and generating blocks.
    /// </summary>
    public interface ILibraryManager
    {
        /// <summary>
        /// Gets the path to the project library.
        /// </summary>
        string ProjectLibPath { get; }

        /// <summary>
        /// Gets or sets the TIA Portal instance to work with.
        /// </summary>
        TiaPortal WorkTiaPortal { get; set; }

        /// <summary>
        /// Opens a global library from the specified path.
        /// </summary>
        /// <param name="libraryPath">Path to the library file.</param>
        /// <returns>The opened UserGlobalLibrary instance.</returns>
        UserGlobalLibrary OpenLibrary(string libraryPath);

        /// <summary>
        /// Searches for a library file using a file dialog with the specified filter.
        /// </summary>
        /// <param name="filter">File filter for the dialog.</param>
        /// <returns>Path to the selected library file or empty string if none selected.</returns>
        string SearchLibrary(string filter);

        /// <summary>
        /// Closes the currently opened global library.
        /// </summary>
        /// <returns>True if the library was closed successfully, false otherwise.</returns>
        bool CloseLibrary();

        /// <summary>
        /// Updates a project library from the global library.
        /// </summary>
        /// <param name="libraryPath">Path to the global library.</param>
        /// <param name="projectLibrary">Project library to update.</param>
        /// <returns>True if the update was successful, false otherwise.</returns>
        bool UpdatePrjLibraryFromGlobal(string libraryPath, ProjectLibrary projectLibrary);

        /// <summary>
        /// Generates a block from the library and adds it to the PLC software.
        /// </summary>
        /// <param name="plcSoftware">Target PLC software instance.</param>
        /// <param name="projectLibrary">Source project library.</param>
        /// <param name="pathName">Path to the block in the library.</param>
        /// <param name="typeName">Name of the block type.</param>
        /// <param name="groupName">Optional group name for the new block.</param>
        /// <returns>True if the block was generated successfully, false otherwise.</returns>
        bool GenerateBlockFromLibrary(PlcSoftware plcSoftware, ProjectLibrary projectLibrary, string pathName, string typeName, string groupName = null);

        /// <summary>
        /// Generates a UDT (User Defined Type) from the library and adds it to the PLC software.
        /// </summary>
        /// <param name="plcSoftware">Target PLC software instance.</param>
        /// <param name="projectLibrary">Source project library.</param>
        /// <param name="typeName">Name of the type.</param>
        /// <param name="blockName">Name of the block to create.</param>
        /// <param name="groupName">Optional group name for the new UDT.</param>
        /// <returns>True if the UDT was generated successfully, false otherwise.</returns>
        bool GenerateUDTFromLibrary(PlcSoftware plcSoftware, ProjectLibrary projectLibrary, string typeName, string blockName, string groupName = null);

        /// <summary>
        /// Cleans up the main library by removing unused types.
        /// </summary>
        /// <param name="projectLibrary">Project library to clean up.</param>
        /// <returns>True if the cleanup was successful, false otherwise.</returns>
        bool CleanUpMainLibrary(ProjectLibrary projectLibrary);
    }
}