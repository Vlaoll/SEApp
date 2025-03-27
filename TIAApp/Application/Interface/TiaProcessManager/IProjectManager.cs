// Ignore Spelling: Conf

using Siemens.Engineering;
using System;

namespace seConfSW.Services
{
    /// <summary>
    /// Interface for managing TIA Portal projects
    /// </summary>
    public interface IProjectManager : IDisposable
    {
        /// <summary>
        /// Gets the path of the current project
        /// </summary>
        string ProjectPath { get; }

        /// <summary>
        /// Gets the active TIA Portal instance
        /// </summary>
        TiaPortal WorkTiaPortal { get; }

        /// <summary>
        /// Gets the active project instance
        /// </summary>
        Project WorkProject { get; }

        /// <summary>
        /// Opens a TIA Portal project
        /// </summary>
        /// <param name="tiaPortal">TIA Portal instance</param>
        /// <param name="projectPath">Path to the project file</param>
        /// <returns>Opened project or null if failed</returns>
        Project OpenProject(TiaPortal tiaPortal, string projectPath);

        /// <summary>
        /// Closes the current project
        /// </summary>
        void CloseProject();

        /// <summary>
        /// Creates a new TIA Portal project
        /// </summary>
        /// <param name="tiaPortal">TIA Portal instance</param>
        /// <param name="projectPath">Path for the new project</param>
        /// <returns>Created project or null if failed</returns>
        Project CreateProject(TiaPortal tiaPortal, string projectPath);

        /// <summary>
        /// Saves the current project
        /// </summary>
        void SaveProject();

        /// <summary>
        /// Searches for a project file using file dialog
        /// </summary>
        /// <param name="filter">File filter string</param>
        /// <returns>Selected project path or empty string</returns>
        string SearchProject(string filter);

        /// <summary>
        /// Starts TIA Portal application
        /// </summary>
        /// <param name="isVisibleTia">Whether to show UI</param>
        /// <returns>True if started successfully</returns>
        bool StartTIA(bool isVisibleTia = false);

        /// <summary>
        /// Connects to a running TIA Portal instance
        /// </summary>
        /// <returns>Connection status</returns>
        ConnectionStatus ConnectToTiaPortal();
    }
}