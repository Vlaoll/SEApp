// Ignore Spelling: Conf
using Microsoft.Win32;
using Serilog;
using System;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Management;
using System.Net.NetworkInformation;
using System.Security.Cryptography;
using System.Text;

namespace seConfSW.Services
{
    /// <summary>
    /// Manages license validation, generation and cryptographic operations.
    /// Handles license file creation, validation against hardware identifiers,
    /// and cryptographic signing/verification.
    /// </summary>
    public class PermitManager
    {
        #region Constants

        /// <summary>
        /// Prefix for all log entries related to licensing
        /// </summary>
        private const string LogPrefix = "[License]";

        /// <summary>
        /// Default validity period (in days) for new licenses
        /// </summary>
        private const string DefaultDaysValid = "60";

        /// <summary>
        /// Windows group name required for TIA Openness access
        /// </summary>
        private const string TiaGroupName = "Siemens TIA Openness";

        #endregion
        #region Properties and Dependencies

        /// <summary>
        /// Logger instance for recording operational events
        /// </summary>
        private readonly ILogger _logger;

        /// <summary>
        /// Configuration service providing settings and paths
        /// </summary>
        private readonly IConfigurationService _configuration;

        /// <summary>
        /// Service handling license data generation
        /// </summary>
        private readonly LicenseDataService _licenseDataService;

        #endregion
        #region Constructor

        /// <summary>
        /// Initializes a new instance of the LicenseManager
        /// </summary>
        /// <param name="configuration">Configuration service</param>
        /// <param name="logger">Logger instance</param>
        /// <exception cref="ArgumentNullException">Thrown when dependencies are null</exception>
        public PermitManager(IConfigurationService configuration, ILogger logger)
        {
            _logger = logger ?? throw new ArgumentNullException(nameof(logger));
            _configuration = configuration ?? throw new ArgumentNullException(nameof(configuration));
            _licenseDataService = new LicenseDataService(logger);
        }

        #endregion
        #region Public Methods

        /// <summary>
        /// Validates the current license file against system hardware and expiration date
        /// </summary>
        /// <returns>True if license is valid, false otherwise</returns>
        public bool CheckLicense()
        {
            try
            {
                _logger.Information($"{LogPrefix} Starting license validation checks");

                // Validate license file existence
                string licenseFilePath = _configuration.LicenseFile;
                _logger.Debug($"{LogPrefix} Checking license file at: {licenseFilePath}");

                if (!File.Exists(licenseFilePath))
                {
                    _logger.Warning($"{LogPrefix} License file not found at {licenseFilePath}. Creating default license.");
                    CreateDefaultLicenseJson();
                    return false;
                }

                // Read and decrypt license content
                string encryptedContent = File.ReadAllText(licenseFilePath);
                string decryptedContent = DecryptLicense(encryptedContent, _configuration.LicenseSalt);
                string[] licenseLines = decryptedContent.Split(
                    new[] { Environment.NewLine },
                    StringSplitOptions.RemoveEmptyEntries);

                // Validate license format
                if (licenseLines.Length < 4)
                {
                    _logger.Warning($"{LogPrefix} Invalid license format. Expected 4 lines, got {licenseLines.Length}");
                    return false;
                }

                // Parse license components
                string licensedMac = licenseLines[0];
                string licensedCpu = licenseLines[1];
                string expiryDateStr = licenseLines[2];
                string signature = licenseLines[3];

                // Get current hardware identifiers
                string currentMac = GetCurrentMacAddress();
                string currentCpu = GetCurrentCpuId();

                // Validate MAC address
                if (string.IsNullOrEmpty(currentMac) || currentMac != licensedMac)
                {
                    _logger.Warning($"{LogPrefix} MAC address mismatch. Licensed: {licensedMac}, Current: {currentMac}");
                    return false;
                }

                // Validate CPU ID
                if (string.IsNullOrEmpty(currentCpu) || currentCpu != licensedCpu)
                {
                    _logger.Warning($"{LogPrefix} CPU ID mismatch. Licensed: {licensedCpu}, Current: {currentCpu}");
                    return false;
                }

                // Validate expiration date
                if (!DateTime.TryParseExact(expiryDateStr, "yyyy-MM-dd",
                    CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime expiryDate))
                {
                    _logger.Warning($"{LogPrefix} Invalid date format: {expiryDateStr}");
                    return false;
                }

                if (expiryDate < DateTime.Now)
                {
                    _logger.Warning($"{LogPrefix} License expired on {expiryDate:yyyy-MM-dd}");
                    return false;
                }

                // Validate cryptographic signature
                string signedData = licensedMac + licensedCpu + expiryDateStr;
                string computedSignature = ComputeSHA256Hash(signedData + _configuration.LicenseSalt);

                if (computedSignature != signature)
                {
                    _logger.Warning($"{LogPrefix} Invalid license signature");
                    return false;
                }

                // Validate Windows group membership
                System.Security.Principal.WindowsPrincipal principal = new System.Security.Principal.WindowsPrincipal(System.Security.Principal.WindowsIdentity.GetCurrent());
                if (!principal.IsInRole(TiaGroupName))
                {
                    _logger.Warning($"{LogPrefix} User must be added to 'Siemens TIA Openness' group");
                    return false;
                }

                // Validate TIA Portal installation
                RegistryKey filePathReg = Registry.LocalMachine.OpenSubKey(_configuration.SiemensRegistryPath);
                if (filePathReg == null)
                {
                    _logger.Warning($"{LogPrefix} TIA Portal Openness is missing or incorrect version installed");
                    return false;
                }

                _logger.Information($"{LogPrefix} All permissions checked successfully. License validation successful");
                return true;
            }
            catch (Exception ex)
            {
                _logger.Error(ex, $"{LogPrefix} License validation failed: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// Generates a new license file from a configuration JSON
        /// </summary>
        /// <returns>True if generation succeeded, false otherwise</returns>
        public bool GenerateLicense()
        {
            try
            {
                _logger.Information($"{LogPrefix} Starting license generation");

                // Prompt user to select configuration
                var dialog = new OpenFileDialog
                {
                    Filter = "JSON files (*.json)|*.json",
                    Multiselect = false,
                    RestoreDirectory = true
                };

                if (dialog.ShowDialog() != true)
                {
                    _logger.Information($"{LogPrefix} License generation canceled by user");
                    return false;
                }

                string settingsPath = dialog.FileName;
                _logger.Debug($"{LogPrefix} Using configuration: {settingsPath}");

                // Parse configuration
                string jsonContent = File.ReadAllText(settingsPath);
                string macAddress = ExtractJsonValue(jsonContent, "MacAddress");
                string cpuId = ExtractJsonValue(jsonContent, "CpuId");
                string licensePath = _configuration.LicenseFile;
                string salt = _configuration.LicenseSalt;
                int validityDays = int.Parse(DefaultDaysValid);

                // Validate required fields
                if (string.IsNullOrEmpty(macAddress) || string.IsNullOrEmpty(cpuId))
                {
                    _logger.Warning($"{LogPrefix} Configuration missing MAC or CPU ID");
                    return false;
                }

                // Generate license data
                var (_, expiryDate) = _licenseDataService.GenerateLicenseData(validityDays);
                string expiryDateStr = expiryDate.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture);

                // Create and sign license content
                string dataToSign = macAddress + cpuId + expiryDateStr;
                string signature = ComputeSHA256Hash(dataToSign + salt);
                string licenseContent = string.Join(Environment.NewLine, macAddress, cpuId, expiryDateStr, signature);

                // Encrypt and save
                string encryptedContent = EncryptLicense(licenseContent, salt);

                // Remove existing file if present
                if (File.Exists(licensePath))
                {
                    File.Delete(licensePath);
                }
                File.WriteAllText(licensePath, encryptedContent);

                _logger.Information($"{LogPrefix} Generated license at {licensePath} valid until {expiryDateStr}");
                return true;
            }
            catch (FormatException ex)
            {
                _logger.Error(ex, $"{LogPrefix} Date format error: {ex.Message}");
                return false;
            }
            catch (CryptographicException ex)
            {
                _logger.Error(ex, $"{LogPrefix} Encryption failed: {ex.Message}");
                return false;
            }
            catch (IOException ex)
            {
                _logger.Error(ex, $"{LogPrefix} File operation failed: {ex.Message}");
                return false;
            }
            catch (Exception ex)
            {
                _logger.Error(ex, $"{LogPrefix} Generation failed: {ex.Message}");
                return false;
            }
        }

        #endregion
        #region Private Helper Methods

        /// <summary>
        /// Creates a default license file with current hardware identifiers
        /// </summary>
        /// <returns>True if creation succeeded, false otherwise</returns>
        private bool CreateDefaultLicenseJson()
        {
            try
            {
                string licenseJsonPath = Path.Combine(
                    AppDomain.CurrentDomain.BaseDirectory,
                    _configuration.LicenseInit);

                _logger.Information($"{LogPrefix} Creating default license at {licenseJsonPath}");

                // Remove existing file if present
                if (File.Exists(licenseJsonPath))
                {
                    File.Delete(licenseJsonPath);
                }

                // Get current hardware info
                string macAddress = GetCurrentMacAddress();
                string cpuId = GetCurrentCpuId();
                string licenseFilePath = _configuration.LicenseFile;

                // Create JSON structure
                string defaultContent = $@"{{
    ""License"":{{
        ""MacAddress"": ""{macAddress}"",
        ""CpuId"": ""{cpuId}"",
        ""FilePath"": ""{licenseFilePath}""}}
}}";

                File.WriteAllText(licenseJsonPath, defaultContent);

                _logger.Information($"{LogPrefix} Created default license for MAC: {macAddress}, CPU: {cpuId}");
                return true;
            }
            catch (Exception ex)
            {
                _logger.Error(ex, $"{LogPrefix} Failed to create default license: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// Extracts a value from JSON using simple string parsing
        /// </summary>
        /// <param name="json">JSON content</param>
        /// <param name="key">Key to extract</param>
        /// <returns>Extracted value or null</returns>
        private string ExtractJsonValue(string json, string key)
        {
            string searchPattern = $"\"{key}\": \"";
            int startIndex = json.IndexOf(searchPattern) + searchPattern.Length;

            if (startIndex < searchPattern.Length)
                return null;

            int endIndex = json.IndexOf("\"", startIndex);
            return endIndex > startIndex ?
                json.Substring(startIndex, endIndex - startIndex) :
                null;
        }

        /// <summary>
        /// Encrypts license content using AES-256
        /// </summary>
        /// <param name="plainText">Content to encrypt</param>
        /// <param name="salt">Encryption salt</param>
        /// <returns>Base64-encoded encrypted content</returns>
        /// <exception cref="ArgumentNullException">Thrown when input parameters are null</exception>
        private string EncryptLicense(string plainText, string salt)
        {
            if (string.IsNullOrEmpty(plainText))
                throw new ArgumentNullException(nameof(plainText));
            if (string.IsNullOrEmpty(salt))
                throw new ArgumentNullException(nameof(salt));

            try
            {
                byte[] key = ComputeSHA256HashBytes(salt);
                byte[] iv = new byte[16];
                Array.Copy(key, iv, 16); // Use first 16 bytes of hash as IV

                using (Aes aes = Aes.Create())
                {
                    aes.Key = key;
                    aes.IV = iv;
                    aes.Padding = PaddingMode.PKCS7;
                    aes.Mode = CipherMode.CBC;

                    using (var encryptor = aes.CreateEncryptor())
                    using (var ms = new MemoryStream())
                    {
                        using (var cs = new CryptoStream(ms, encryptor, CryptoStreamMode.Write))
                        using (var sw = new StreamWriter(cs))
                        {
                            sw.Write(plainText);
                        }
                        return Convert.ToBase64String(ms.ToArray());
                    }
                }
            }
            catch (CryptographicException ex)
            {
                _logger.Error(ex, $"{LogPrefix} Encryption failed: {ex.Message}");
                throw;
            }
        }

        /// <summary>
        /// Decrypts license content using AES-256
        /// </summary>
        /// <param name="cipherText">Encrypted content to decrypt</param>
        /// <param name="salt">Decryption salt</param>
        /// <returns>Decrypted plain text</returns>
        /// <exception cref="ArgumentNullException">Thrown when input parameters are null</exception>
        private string DecryptLicense(string cipherText, string salt)
        {
            if (string.IsNullOrEmpty(cipherText))
                throw new ArgumentNullException(nameof(cipherText));
            if (string.IsNullOrEmpty(salt))
                throw new ArgumentNullException(nameof(salt));

            try
            {
                byte[] key = ComputeSHA256HashBytes(salt);
                byte[] iv = new byte[16];
                Array.Copy(key, iv, 16);

                byte[] cipherBytes = Convert.FromBase64String(cipherText);

                using (Aes aes = Aes.Create())
                {
                    aes.Key = key;
                    aes.IV = iv;
                    aes.Padding = PaddingMode.PKCS7;
                    aes.Mode = CipherMode.CBC;

                    using (var decryptor = aes.CreateDecryptor())
                    using (var ms = new MemoryStream(cipherBytes))
                    using (var cs = new CryptoStream(ms, decryptor, CryptoStreamMode.Read))
                    using (var sr = new StreamReader(cs))
                    {
                        return sr.ReadToEnd();
                    }
                }
            }
            catch (FormatException ex)
            {
                _logger.Error(ex, $"{LogPrefix} Invalid Base64 data: {ex.Message}");
                throw;
            }
            catch (CryptographicException ex)
            {
                _logger.Error(ex, $"{LogPrefix} Decryption failed: {ex.Message}");
                throw;
            }
        }

        /// <summary>
        /// Computes SHA-256 hash bytes from input string
        /// </summary>
        /// <param name="input">String to hash</param>
        /// <returns>Byte array containing hash</returns>
        /// <exception cref="ArgumentNullException">Thrown when input is null</exception>
        private byte[] ComputeSHA256HashBytes(string input)
        {
            if (input == null)
                throw new ArgumentNullException(nameof(input));

            using (SHA256 sha256 = SHA256.Create())
            {
                return sha256.ComputeHash(Encoding.UTF8.GetBytes(input));
            }
        }

        /// <summary>
        /// Computes SHA-256 hash string from input string
        /// </summary>
        /// <param name="input">String to hash</param>
        /// <returns>Hexadecimal representation of hash</returns>
        private string ComputeSHA256Hash(string input)
        {
            byte[] hashBytes = ComputeSHA256HashBytes(input);
            var builder = new StringBuilder(hashBytes.Length * 2);

            foreach (byte b in hashBytes)
            {
                builder.Append(b.ToString("x2"));
            }

            return builder.ToString();
        }

        /// <summary>
        /// Retrieves MAC address of primary network interface
        /// </summary>
        /// <returns>MAC address string or empty string on failure</returns>
        private string GetCurrentMacAddress()
        {
            try
            {
                return NetworkInterface.GetAllNetworkInterfaces()
                    .Where(nic => nic.OperationalStatus == OperationalStatus.Up)
                    .Select(nic => nic.GetPhysicalAddress().ToString())
                    .FirstOrDefault() ?? string.Empty;
            }
            catch (Exception ex)
            {
                _logger.Error(ex, $"{LogPrefix} MAC address retrieval failed: {ex.Message}");
                return string.Empty;
            }
        }

        /// <summary>
        /// Retrieves CPU ID using WMI
        /// </summary>
        /// <returns>CPU ID string or empty string on failure</returns>
        private string GetCurrentCpuId()
        {
            try
            {
                using (var searcher = new ManagementObjectSearcher("SELECT ProcessorId FROM Win32_Processor"))
                {
                    return searcher.Get().Cast<ManagementObject>()
                        .Select(mo => mo["ProcessorId"]?.ToString())
                        .FirstOrDefault() ?? string.Empty;
                }
            }
            catch (Exception ex)
            {
                _logger.Error(ex, $"{LogPrefix} CPU ID retrieval failed: {ex.Message}");
                return string.Empty;
            }
        }

        #endregion
    }
}