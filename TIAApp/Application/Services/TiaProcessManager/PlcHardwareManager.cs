// Ignore Spelling: Plc Conf

using System;
using System.Linq;
using Siemens.Engineering;
using Serilog;
using Siemens.Engineering.SW;
using Siemens.Engineering.HW.Features;

namespace seConfSW.Services
{    
    /// <summary>
    /// Provides functionality to manage Hardware configuration in the TIA Portal.
    /// Handles device creation and PLC software retrieval operations.
    /// </summary>
    public class PlcHardwareManager : IPlcHardwareManager
    {
        #region Constants
        /// <summary>
        /// Prefix for log messages from this class
        /// </summary>
        private const string LogPrefix = "[TIA/HwM]";
        #endregion
        #region Properties
        /// <summary>
        /// Logger instance for recording operations
        /// </summary>
        private readonly ILogger _logger;
        #endregion
        #region Constructor
        /// <summary>
        /// Initializes a new instance of the <see cref="PlcHardwareManager"/> class.
        /// </summary>
        /// <param name="logger">The logger instance for recording operations</param>
        /// <exception cref="ArgumentNullException">Thrown when logger is null</exception>
        public PlcHardwareManager(ILogger logger)
        {
            _logger = logger ?? throw new ArgumentNullException(nameof(logger));
        }
        #endregion
        #region Public Methods
        /// <inheritdoc />
        public bool AddHW(Project project, string nameDevice, string orderNo, string version)
        {
            // Validate input parameters
            if (project == null) throw new ArgumentNullException(nameof(project));
            if (string.IsNullOrEmpty(nameDevice)) throw new ArgumentException("Device name cannot be null or empty.", nameof(nameDevice));
            if (string.IsNullOrEmpty(orderNo)) throw new ArgumentException("Order number cannot be null or empty.", nameof(orderNo));
            if (string.IsNullOrEmpty(version)) throw new ArgumentException("Version cannot be null or empty.", nameof(version));

            try
            {
                // Construct the MLFB (Order Number) string
                string mlfb = $"OrderNumber:{orderNo}/{version}";
                string devName = "station" + nameDevice;

                // Check if the device already exists
                bool deviceExists = project.Devices.Any(d => d.DeviceItems.Any(item => item.Name == devName) || d.Name == devName);

                if (deviceExists)
                {
                    _logger.Information($"{LogPrefix} Device already exists: {nameDevice}");
                    return false;
                }

                // Create the new device
                var device = project.Devices.CreateWithItem(mlfb, nameDevice, devName);
                _logger.Information($"{LogPrefix} Added device: {nameDevice} with Order Number: {orderNo} and Version: {version}");

                return true;
            }
            catch (Exception ex)
            {
                _logger.Error(ex, $"{LogPrefix} Error occurred while adding hardware device: {nameDevice}");
                return false;
            }
        }

        /// <inheritdoc />
        public PlcSoftware GetPLC(Project project, string plcName)
        {
            // Validate input parameters
            if (project == null) throw new ArgumentNullException(nameof(project));
            if (string.IsNullOrEmpty(plcName)) throw new ArgumentException("PLC name cannot be null or empty.", nameof(plcName));

            try
            {
                // Iterate through all devices in the project
                foreach (var device in project.Devices)
                {
                    // Find the device item that matches the PLC name
                    var deviceItem = device.DeviceItems.FirstOrDefault(item => item.Name.Contains(plcName));
                    if (deviceItem != null)
                    {
                        // Get the SoftwareContainer service from the device item
                        var softwareContainer = deviceItem.GetService<SoftwareContainer>();
                        if (softwareContainer != null)
                        {
                            // Return the PLC software instance
                            _logger.Information($"{LogPrefix} Found PLC: {plcName}");
                            return softwareContainer.Software as PlcSoftware;
                        }
                    }
                }

                _logger.Warning($"{LogPrefix} PLC not found: {plcName}");
                return null;
            }
            catch (Exception ex)
            {
                _logger.Error(ex, $"{LogPrefix} Error occurred while retrieving PLC: {plcName}");
                return null;
            }
        }
        #endregion
        #region Private Helper Methods
        // Add any private helper methods here in future if needed
        #endregion
    }
}