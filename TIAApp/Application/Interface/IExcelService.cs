using System;
using System.Collections.Generic;

namespace seConfSW.Services
{
    public interface IExcelService
    {        
        event EventHandler<string> MessageUpdated;

        /// <summary>
        /// Creates an Excel database by reading and processing the configured Excel file.
        /// </summary>
        /// <returns>True if the operation succeeded, false otherwise</returns>
        bool CreateExcelDB();
        /// <summary>
        /// Retrieves the PLC data structure from the Excel reader.
        /// </summary>
        /// <returns>List of PLC data blocks</returns>
        List<Domain.Models.dataPLC> GetExcelDataReader();
        /// <summary>
        /// Closes the currently open Excel file.
        /// </summary>
        void CloseExcelFile();
    }
}