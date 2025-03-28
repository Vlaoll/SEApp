// Ignore Spelling: Conf

namespace seConfSW.Services
{
    /// <summary>
    /// Provides configuration settings for the TIA Portal process.
    /// </summary>
    public interface IConfigurationService
    {
        /// <summary>
        /// Gets the Siemens Registry Path from the configuration.
        /// </summary>
        string SiemensRegistryPath { get; }
       

        /// <summary>
        /// Gets the project file filter for the TIA Portal.
        /// </summary>
        string ProjectFilter { get; }

        /// <summary>
        /// Gets the library file filter for the TIA Portal.
        /// </summary>
        string LibraryFilter { get; }        

        /// <summary>
        /// Gets the default source path for the TIA Portal.
        /// </summary>
        string DefaultSourcePath { get; }

        /// <summary>
        /// Gets the export path for the TIA Portal.
        /// </summary>
        string ExportPath { get; }

        /// <summary>
        /// Gets the source database path for the TIA Portal.
        /// </summary>
        string SourceDBPath { get; }

        /// <summary>
        /// Gets the source tag path for the TIA Portal.
        /// </summary>
        string SourceTagPath { get; }

        /// <summary>
        /// Gets the default project path for the TIA Portal.
        /// </summary>
        string DefaultProjectPath { get; }

        /// <summary>
        /// Gets the template path for the TIA Portal.
        /// </summary>
        string TemplatePath { get; }

        /// <summary>
        /// Gets the excel file filter for data collect.
        /// </summary>
        string ExcelFilter { get; }

        /// <summary>
        /// Gets the excel main sheet name.
        /// </summary>
        string MainExcelSheetName { get; }

        /// <summary>
        /// Gets the license file.
        /// </summary>
        string LicenseFile { get; }

        /// <summary>
        /// Gets the init license file.
        /// </summary>
        string LicenseInit { get; }

        // <summary>
        /// Gets the license salt.
        /// </summary>
        string LicenseSalt { get; }

        /// <summary>
        /// Gets a value indicating whether the TIA Portal is visible.
        /// </summary>
        bool IsVisibleTia { get; }
    }
}