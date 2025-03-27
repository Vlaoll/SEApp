// Ignore Spelling: Conf

using System;
using Siemens.Engineering.Compiler;
using Siemens.Engineering.HW;
using Siemens.Engineering.HW.Features;
using Siemens.Engineering.SW;
using Siemens.Engineering.Hmi;
using Serilog;
using Siemens.Engineering;

namespace seConfSW.Services
{
    

    /// <summary>
    /// Provides services for compiling devices in a TIA project.
    /// Handles compilation of PLC software and HMI targets with proper logging.
    /// </summary>
    public class CompilerManager : ICompilerManager
    {
        #region Constants

        /// <summary>
        /// Prefix for all log messages from this class.
        /// </summary>
        private const string LogPrefix = "[TIA/CM]";

        #endregion
        #region Properties

        /// <summary>
        /// Logger instance for recording compilation events and errors.
        /// </summary>
        private readonly ILogger _logger;

        #endregion
        #region Constructor

        /// <summary>
        /// Initializes a new instance of the <see cref="CompilerManager"/> class.
        /// </summary>        
        /// <param name="logger">The Serilog logger instance for logging messages.</param>
        /// <exception cref="ArgumentNullException">Thrown if <paramref name="logger"/> is null.</exception>
        public CompilerManager(ILogger logger)
        {
            _logger = logger ?? throw new ArgumentNullException(nameof(logger));
        }

        #endregion
        #region Public Methods

        /// <inheritdoc />
        public void Compile(string devName, Project project)
        {
            if (string.IsNullOrEmpty(devName))
            {
                throw new ArgumentNullException(nameof(devName));
            }

            bool deviceFound = false;

            try
            {
                // Iterate through all devices in the project
                foreach (var device in project.Devices)
                {
                    var deviceItemAggregation = device.DeviceItems;

                    // Check each device item for matching name
                    foreach (var deviceItem in deviceItemAggregation)
                    {
                        if (deviceItem.Name == devName || device.Name == devName)
                        {
                            var softwareContainer = deviceItem.GetService<SoftwareContainer>();
                            if (softwareContainer != null)
                            {
                                // Handle PLC Software compilation
                                if (softwareContainer.Software is PlcSoftware plcSoftware)
                                {
                                    deviceFound = true;
                                    CompileSoftware(plcSoftware);
                                }
                                // Handle HMI Target compilation
                                else if (softwareContainer.Software is HmiTarget hmiTarget)
                                {
                                    deviceFound = true;
                                    CompileSoftware(hmiTarget);
                                }
                            }
                        }
                    }
                }

                if (!deviceFound)
                {
                    _logger.Information($"{LogPrefix} Didn't find device with name {devName}");
                }
            }
            catch (Exception ex)
            {
                _logger.Error(ex, $"{LogPrefix} An error occurred during compilation: {ex.Message}");
                throw; // Re-throw the exception to preserve stack trace
            }
        }

        #endregion
        #region Private Helper Methods

        /// <summary>
        /// Compiles the specified software component and logs the results.
        /// </summary>
        /// <param name="compilableSoftware">The software component to compile (PLC or HMI).</param>
        /// <remarks>
        /// Handles both PLC software and HMI targets, logging compilation progress and results.
        /// </remarks>
        private void CompileSoftware(IEngineeringServiceProvider compilableSoftware)
        {
            var compiler = compilableSoftware.GetService<ICompilable>();
            var software = (Software)compilableSoftware;

            if (compiler == null)
            {
                _logger.Warning($"{LogPrefix} Compiler service not found for {software.Name}");
                return;
            }

            try
            {
                _logger.Information($"{LogPrefix} Compiling {software.Name}");
                var result = compiler.Compile();

                // Log compilation results
                _logger.Information(
                    $"{LogPrefix} Compiling {software.Name}: " +
                    $"State: {result.State} / " +
                    $"Warning Count: {result.WarningCount} / " +
                    $"Error Count: {result.ErrorCount}");
            }
            catch (Exception ex)
            {
                _logger.Error(ex, $"{LogPrefix} Failed to compile {software.Name}");
                throw;
            }
        }

        #endregion
    }
}