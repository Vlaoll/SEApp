// Ignore Spelling: Api Conf

using Microsoft.Extensions.Configuration;
using System;

namespace seConfSW.Services
{
    /// <summary>
    /// Configuration class that provides settings for the TIA Portal process.
    /// Implements the <see cref="IConfigurationService"/> interface.
    /// </summary>
    public class Configuration : IConfigurationService
    {
        private readonly IConfiguration _configuration;

        /// <summary>
        /// Initializes a new instance of the <see cref="Configuration"/> class.
        /// </summary>
        /// <param name="configuration">The configuration object.</param>
        /// <exception cref="ArgumentNullException">Thrown when the configuration object is null.</exception>
        public Configuration(IConfiguration configuration)
        {
            _configuration = configuration ?? throw new ArgumentNullException(nameof(configuration));
        }

        /// <inheritdoc/>
        public string SiemensRegistryPath => _configuration["TiaPortal:SiemensRegistryPathKey"] ?? "SOFTWARE\\Siemens\\Automation\\Openness";

        /// <inheritdoc/>
        public string SiemensApiPath => _configuration["TiaPortal:SiemensApiPath"] ?? "C:\\Program Files\\Siemens\\Automation\\Portal V19\\PublicAPI\\V19\\";

        /// <inheritdoc/>
        public string ProjectFilter => _configuration["TiaPortal:Filter"] ?? "TIA Portal |*.ap20; *.ap19; *.ap18; *.ap17";

        /// <inheritdoc/>
        public string LibraryFilter => _configuration["TiaLibrary:Filter"] ?? "TIA Library |*.al20; *.al19; *.al18; *.al17";
        

        /// <inheritdoc/>
        public string DefaultSourcePath => _configuration["TiaPortal:DefaultSourcePath"] ?? "samples\\source\\";

        /// <inheritdoc/>
        public string ExportPath => _configuration["TiaPortal:ExportPath"] ?? "Samples\\export\\";

        /// <inheritdoc/>
        public string SourceDBPath => _configuration["TiaPortal:SourceDBPath"] ?? "Samples\\sourceDB\\";

        /// <inheritdoc/>
        public string SourceTagPath => _configuration["TiaPortal:SourceTagPath"] ?? "samples\\tag\\";

        /// <inheritdoc/>
        public string DefaultProjectPath => _configuration["TiaPortal:DefaultProjectPath"] ?? "Samples\\project\\";

        /// <inheritdoc/>
        public string TemplatePath => _configuration["TiaPortal:TemplatePath"] ?? "samples\\template\\";

        /// <inheritdoc/>
        public string ExcelFilter => _configuration["Excel:Filter"] ?? "Excel |*.xlsx";

        /// <inheritdoc/>
        public string MainExcelSheetName => _configuration["Excel:MainSheetName"] ?? "Main";

        /// <inheritdoc/>
        public string LicenseFile => _configuration["License:LicenseFile"] ?? "license.lic";
        public string LicenseInit => _configuration["License:LicenseInit"] ?? "license.json";

        /// <inheritdoc/>
        public string LicenseSalt => "Se2847!!";

        /// <inheritdoc/>
        public bool IsVisibleTia => _configuration.GetValue<bool>("TiaPortal:IsVisibleTia", false);








        
    }
}