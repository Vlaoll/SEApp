// Ignore Spelling: Conf

using seConfSW.Domain.Models;
using System.Collections.Generic;

namespace seConfSW
{
    /// <summary>
    /// Defines the contract for reading and processing Excel data related to PLC configurations.
    /// </summary>
    public interface IExcelDataReader
    {
        /// <summary>
        /// Gets the last message generated during Excel processing, such as success or error details.
        /// </summary>
        string Message { get; }

        /// <summary>
        /// Gets the list of PLC data structures populated from Excel.
        /// </summary>
        List<dataPLC> BlocksStruct { get; }

        /// <summary>
        /// Opens a file dialog to select an Excel project file and returns its path.
        /// </summary>
        /// <param name="filter">The file filter for the dialog (e.g., "Excel |*.xlsx;*.xlsm"). Defaults to "Excel |*.xlsx;*.xlsm".</param>
        /// <returns>The selected file path or an empty string if no file is selected.</returns>
        string SearchProject(string filter = "Excel |*.xlsx;*.xlsm");

        /// <summary>
        /// Opens an Excel file and initializes worksheets for processing.
        /// </summary>
        /// <param name="filename">The path to the Excel file to open.</param>
        /// <param name="mainSheetName">The name of the main worksheet to load. Defaults to "Main".</param>
        /// <returns>True if the file is successfully opened, false otherwise.</returns>
        bool OpenExcelFile(string filename, string mainSheetName = "Main");

        /// <summary>
        /// Closes the currently open Excel file.
        /// </summary>
        /// <param name="save">Indicates whether to save changes before closing. Defaults to false.</param>
        /// <returns>True if the file is successfully closed, false otherwise.</returns>
        bool CloseExcelFile(bool save = false);

        /// <summary>
        /// Reads and processes object data from the main worksheet based on a status filter.
        /// </summary>
        /// <param name="status">The status value to filter rows in the main worksheet.</param>
        /// <param name="maxInstanceCount">The maximum number of instances allowed per PLC. Defaults to 250.</param>
        /// <returns>True if at least one row was processed successfully, false otherwise.</returns>
        bool ReadExcelObjectData(string status, int maxInstanceCount = 250);

        /// <summary>
        /// Reads and processes extended data from a specified worksheet (e.g., PLCData).
        /// </summary>
        /// <param name="sheetBlockDataName">The name of the worksheet containing extended data. Defaults to "PLCData".</param>
        /// <returns>True if the data is successfully processed, false otherwise.</returns>
        bool ReadExcelExtendedData(string sheetBlockDataName = "PLCData");
    }
}