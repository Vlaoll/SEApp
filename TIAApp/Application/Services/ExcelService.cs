// Ignore Spelling: plc Conf Eq Prj excelprj

using seConfSW.Domain.Models;
using Serilog;
using System;
using System.Collections.Generic;

namespace seConfSW.Services
{
    /// <summary>
    /// Service for handling Excel file operations including reading, processing, and managing Excel-based data.
    /// Provides functionality to create an Excel database, retrieve data, and manage file operations.
    /// </summary>
    public class ExcelService : IExcelService
    {
        #region Constants

        /// <summary>
        /// Prefix used for logging messages from this service
        /// </summary>
        private const string LogPrefix = "[Excel]";

        #endregion
        #region Fields

        private readonly ILogger _logger;
        private readonly IConfigurationService _configuration;
        private readonly IExcelDataReader _excelprj;
        private string _excelPath;

        #endregion
        #region Events

        /// <summary>
        /// Event that is triggered when a message needs to be updated
        /// </summary>
        public event EventHandler<string> MessageUpdated;

        #endregion
        #region Properties

        /// <summary>
        /// Gets the data reader instance for accessing Excel data
        /// </summary>
        public IExcelDataReader DataReader => _excelprj;

        #endregion
        #region Constructor

        /// <summary>
        /// Initializes a new instance of the ExcelService class with required dependencies
        /// </summary>
        /// <param name="logger">Logger instance for logging messages</param>
        /// <param name="configuration">Configuration service instance</param>
        /// <param name="excelprj">Excel data reader instance</param>
        /// <exception cref="ArgumentNullException">Thrown when any parameter is null</exception>
        public ExcelService(ILogger logger, IConfigurationService configuration, IExcelDataReader excelprj)
        {
            _configuration = configuration ?? throw new ArgumentNullException(nameof(configuration));
            _logger = logger ?? throw new ArgumentNullException(nameof(logger));
            _excelprj = excelprj ?? throw new ArgumentNullException(nameof(excelprj));
        }

        #endregion
        #region Public Methods

        /// <inheritdoc />
        public bool CreateExcelDB()
        {
            try
            {
                _logger.Information($"{LogPrefix} Starting generation of Excel database");
                _logger.Information($"{LogPrefix} -------------------------------------------------------------------------------------");

                if (string.IsNullOrEmpty(_excelPath))
                {
                    var filter = _configuration.ExcelFilter;
                    _logger.Information($"{LogPrefix} Searching for Excel project file");
                    _excelPath = _excelprj.SearchProject(filter: filter);
                    _logger.Information($"{LogPrefix} Selected Excel file path: {_excelPath}");
                }

                var mainSheetName = _configuration.MainExcelSheetName;
                _logger.Information($"{LogPrefix} Opening Excel file {_excelPath} with main sheet: {mainSheetName}");
                _excelprj.OpenExcelFile(_excelPath, mainSheetName: mainSheetName);

                _logger.Information($"{LogPrefix} Reading object data from Excel file");
                if (!_excelprj.ReadExcelObjectData("Block", 250))
                {
                    _logger.Error($"{LogPrefix} Failed to read object data from Excel file due to incorrect settings");
                    return false;
                }

                _logger.Information($"{LogPrefix} Reading extended data from Excel file");
                if (!_excelprj.ReadExcelExtendedData())
                {
                    _logger.Error($"{LogPrefix} Failed to read extended data from Excel file due to incorrect settings");
                    return false;
                }

                _logger.Information($"{LogPrefix} Closing Excel file");
                _excelprj.CloseExcelFile();
                _logger.Information($"{LogPrefix} Successfully generated Excel database from {_excelPath}");
                _logger.Information($"{LogPrefix} -------------------------------------------------------------------------------------");

                return true;
            }
            catch (Exception ex)
            {
                _logger.Error($"{LogPrefix} Failed to generate Excel database: {ex.Message}");
                return false;
            }
        }

        /// <inheritdoc />
        public List<dataPLC> GetExcelDataReader()
        {
            return _excelprj.BlocksStruct;
        }

        /// <inheritdoc />
        public void CloseExcelFile()
        {
            try
            {
                _logger.Information($"{LogPrefix} Closing Excel file");
                _excelprj?.CloseExcelFile();
                _logger.Information($"{LogPrefix} Excel file closed");
            }
            catch (Exception ex)
            {
                _logger.Error($"{LogPrefix} Failed to close Excel file: {ex.Message}");
            }
        }

        #endregion
        #region Private Helper Methods
        // Currently no private helper methods exist in this class
        // Future private methods should be placed in this region
        #endregion
    }
}