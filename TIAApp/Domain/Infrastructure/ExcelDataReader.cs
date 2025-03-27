// Ignore Spelling: Conf

using ClosedXML.Excel;
using Microsoft.Win32;
using seConfSW.Domain.Models;
using Serilog;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace seConfSW
{
    /// <summary>
    /// Implements functionality to read and process Excel files containing PLC configuration data.
    /// </summary>
    public class ExcelDataReader : IExcelDataReader
    {
        // Fields
        private string excelPath = string.Empty; // Path to the currently open Excel file
        private XLWorkbook workbook; // Workbook instance for the open Excel file
        private IXLWorksheet mainWorkSheet; // Main worksheet from the Excel file
        private List<IXLWorksheet> worksheets; // List of all worksheets in the Excel file
        private readonly ILogger _logger; // Logger instance for logging operations
        private const string LogPrefix = "[Excel]"; // Prefix used in log messages
        private readonly List<dataPLC> blocksStruct = new List<dataPLC>(); // Collection of PLC data structures
        private readonly (int Row, int Col) typeIsExtendedCell = (2, 5); // Cell coordinates to check if equipment is extended

        // Dictionaries for field mappings and coordinates
        private readonly Dictionary<string, int> fieldMain = new Dictionary<string, int>
        {
            {"Status", 1}, {"PLCName", 2}, {"PrjNum", 3}, {"EqName", 4}, {"EqArea", 5}, {"EqComments", 6},
            {"EqType", 7}, {"PLCNumber", 8}, {"InstanceOfName", 9}, {"GroupDB", 10}, {"NameFC", 11},
            {"Variant", 12}, {"ObjName", 13}, {"PicNum", 14}, {"ObjTagName", 15}, {"HMIType", 16},
            {"TypicalPDL", 17}, {"WorkPDL", 18}, {"X", 19}, {"Y", 20}, {"Width", 21}, {"Height", 22},
            {"ScaleMsg", 23}, {"Addition_01", 24}, {"Addition_02", 25}, {"Addition_03", 26},
            {"Addition_04", 27}, {"Addition_05", 28}, {"Addition_06", 29}, {"Addition_07", 30},
            {"Addition_08", 31}, {"Addition_09", 32}, {"Addition_10", 33}
        };
        private readonly Dictionary<string, int> fieldExtendedBlockData = new Dictionary<string, int>
        {
            {"Block Name", 2}, {"Block Number", 3}, {"Block Group", 4}, {"PLCName", 5}, {"Block Type", 6}
        };
        private readonly Dictionary<string, int> fieldExtendedSupportDB = new Dictionary<string, int>
        {
            {"Block Name", 2}, {"Block Group", 3}, {"PLCName", 4}, {"Block Type", 5}, {"Path", 6}, {"Types", 7}
        };
        private readonly Dictionary<string, int> fieldExtendedUserConstant = new Dictionary<string, int>
        {
            {"Block Name", 2}, {"Block Type", 3}, {"Value", 4}, {"PLCName", 5}
        };
        private readonly Dictionary<string, int> fieldExtendedExtFCValue = new Dictionary<string, int>
        {
            {"Name", 2}, {"Type", 3}, {"IO", 4}, {"Comments", 5}, {"FC", 6}, {"PLCName", 7}
        };
        private readonly Dictionary<string, int> fieldTypeFcColumns = new Dictionary<string, int>
        {
            {"name", 2}, {"group", 3}, {"path", 4}, {"isType", 5}
        };
        private readonly Dictionary<string, int> fieldTypeTagColumns = new Dictionary<string, int>
        {
            {"name", 2}, {"link", 3}, {"type", 4}, {"table", 5}, {"comment", 6}, {"variant", 7}
        };
        private readonly Dictionary<string, int> fieldTypeParameterColumns = new Dictionary<string, int>
        {
            {"name", 2}, {"type", 3}, {"link", 4}
        };
        private readonly Dictionary<string, int> fieldTypeExtSupportBlockColumns = new Dictionary<string, int>
        {
            {"name", 2}, {"instanceOfName", 3}, {"number", 4}, {"variant", 5}
        };
        private readonly Dictionary<string, int> fieldTypeConstantColumns = new Dictionary<string, int>
        {
            {"name", 2}, {"type", 3}, {"value", 4}, {"table", 5}
        };
        private readonly Dictionary<string, int> fieldTypeSupportBDColumns = new Dictionary<string, int>
        {
            {"name", 2}, {"number", 3}, {"group", 4}, {"path", 5}, {"isType", 6}, {"isRetain", 7}, {"isOptimazed", 8}
        };
        private readonly Dictionary<string, int> fieldTypeDataBlockValueColumns = new Dictionary<string, int>
        {
            {"name", 2}, {"type", 3}, {"DB", 4}
        };
        private readonly Dictionary<string, (int Row, int StartCol)> typeCoordinates = new Dictionary<string, (int, int)>
        {
            {"FB", (10, 3)}, {"DataTag", (11, 3)}, {"DataParameter", (12, 3)}, {"ExtSupportBlock", (13, 3)},
            {"Constant", (14, 3)}, {"SupportBD", (15, 3)}, {"DataBlockValue", (16, 3)}
        };
        private readonly Dictionary<string, (int Row, int StartCol)> extendedCoordinates = new Dictionary<string, (int, int)>
        {
            {"FCs", (3, 3)},
            {"SupportBDs", (4, 3)},
            {"UserConstants", (5, 3)},
            {"ExtFCValues", (6, 3)}
        };

        // Properties    
        /// <summary>
        /// Gets the list of PLC data structures populated from Excel.
        /// </summary>
        public List<dataPLC> BlocksStruct => blocksStruct;

        /// <summary>
        /// Event raised when a progress message is generated during Excel processing.
        /// </summary>
        public event EventHandler<string> MessageUpdated;

        // Constructor
        /// <summary>
        /// Initializes a new instance of the <see cref="ExcelDataReader"/> class.
        /// </summary>
        /// <param name="logger">Optional logger instance. If null, a default logger is created.</param>
        public ExcelDataReader(ILogger logger)
        {
            _logger = logger ?? throw new ArgumentNullException(nameof(logger)); // Initialize logger
            worksheets = new List<IXLWorksheet>(); // Initialize worksheet list           
        }

        // Public Methods
        /// <summary>
        /// Opens a file dialog to select an Excel project file and returns its path.
        /// </summary>
        /// <param name="filter">The file filter for the dialog (e.g., "Excel |*.xlsx;*.xlsm"). Defaults to "Excel |*.xlsx;*.xlsm".</param>
        /// <returns>The selected file path or an empty string if no file is selected.</returns>
        public string SearchProject(string filter = "Excel |*.xlsx;*.xlsm")
        {
            // Configure and show file dialog
            OpenFileDialog fileSearch = new OpenFileDialog
            {
                Multiselect = false,
                ValidateNames = true,
                DereferenceLinks = false,
                Filter = filter,
                RestoreDirectory = true,
                InitialDirectory = Environment.CurrentDirectory
            };
            fileSearch.ShowDialog();
            string projectPath = fileSearch.FileName;

            // Log file selection if a path is chosen
            if (!string.IsNullOrEmpty(projectPath))
            {
                _logger.Information($"{LogPrefix} Opening excel file {projectPath}");
            }
            return projectPath;
        }
        /// <summary>
        /// Opens an Excel file and initializes worksheets for processing.
        /// </summary>
        /// <param name="filename">The path to the Excel file to open.</param>
        /// <param name="mainSheetName">The name of the main worksheet to load. Defaults to "Main".</param>
        /// <returns>True if the file is successfully opened, false otherwise.</returns>
        public bool OpenExcelFile(string filename, string mainSheetName = "Main")
        {
            try
            {
                excelPath = new FileInfo(filename).FullName; // Get full file path
                workbook = new XLWorkbook(excelPath); // Open the Excel workbook
                worksheets = workbook.Worksheets.ToList(); // Load all worksheets
                mainWorkSheet = worksheets.FirstOrDefault(ws => ws.Name == mainSheetName); // Find main worksheet

                // Check if main worksheet exists
                if (mainWorkSheet == null)
                {
                    _logger.Warning($"{LogPrefix} Main sheet '{mainSheetName}' not found in {excelPath}");
                    return false;
                }

                _logger.Information($"{LogPrefix} Master Excel {excelPath} is opened");
                return true;
            }
            catch (Exception ex)
            {
                _logger.Error($"{LogPrefix} Error while opening excel: {ex.Message}");
                return false;
            }
        }
        /// <summary>
        /// Closes the currently open Excel file.
        /// </summary>
        /// <param name="save">Indicates whether to save changes before closing. Defaults to false.</param>
        /// <returns>True if the file is successfully closed, false otherwise.</returns>
        public bool CloseExcelFile(bool save = false)
        {
            // Check if a workbook is open
            if (workbook == null)
            {
                _logger.Warning($"{LogPrefix} No open workbook to close");
                return false;
            }

            try
            {
                if (save) workbook.Save(); // Save changes if requested
                workbook.Dispose(); // Release workbook resources
                workbook = null; // Clear reference
                _logger.Information($"{LogPrefix} Master Excel {excelPath} is closed");
                return true;
            }
            catch (Exception ex)
            {
                _logger.Error($"{LogPrefix} Error while closing excel: {ex.Message}");
                return false;
            }
        }
        /// <summary>
        /// Reads and processes object data from the main worksheet based on a status filter.
        /// </summary>
        /// <param name="status">The status value to filter rows in the main worksheet.</param>
        /// <param name="maxInstanceCount">The maximum number of instances allowed per PLC. Defaults to 250.</param>
        /// <returns>True if at least one row was processed successfully, false otherwise.</returns>
        public bool ReadExcelObjectData(string status, int maxInstanceCount = 250)
        {
            // Validate main worksheet
            if (mainWorkSheet == null)
            {
                _logger.Warning($"{LogPrefix} Main worksheet is not initialized");
                return false;
            }

            try
            {
                var last = mainWorkSheet.LastCellUsed(); // Get the last used cell
                bool isAnySuccess = false; // Track if any row was processed successfully

                // Iterate through rows starting from row 3
                for (int row = 3; row <= last.Address.RowNumber; row++)
                {
                    // Skip rows not matching the status
                    if (!mainWorkSheet.Cell(row, fieldMain["Status"]).GetValue<string>().Contains(status))
                        continue;

                    try
                    {
                        dataPLC tempItem = InitializePLC(row); // Initialize or retrieve PLC structure
                        // Skip if instance limit is reached
                        if (maxInstanceCount > 0 && tempItem.instanceDB.Count >= maxInstanceCount)
                            continue;

                        // Get equipment worksheet
                        if (!GetEquipmentSheet(row, out string typeEq, out string eqName, out IXLWorksheet eqWorkSheet))
                            continue;

                        ProcessEquipmentData(eqWorkSheet, tempItem, typeEq, eqName); // Process equipment data
                        if (AddFullObjectData(eqWorkSheet, tempItem, row)) // Add instance data
                            isAnySuccess = true;
                    }
                    catch (Exception ex)
                    {
                        _logger.Error($"{LogPrefix} Error processing row {row}: {ex.Message}");
                    }
                }
                _logger.Information($"{LogPrefix} Finished reading object data");
                return isAnySuccess;
            }
            catch (Exception ex)
            {
                _logger.Error($"{LogPrefix} Critical error in excel data processing: {ex.Message}");
                return false;
            }
        }
        /// <summary>
        /// Reads and processes extended data from a specified worksheet (e.g., PLCData).
        /// </summary>
        /// <param name="sheetBlockDataName">The name of the worksheet containing extended data. Defaults to "PLCData".</param>
        /// <returns>True if the data is successfully processed, false otherwise.</returns>
        public bool ReadExcelExtendedData(string sheetBlockDataName = "PLCData")
        {
            try
            {
                // Find the specified worksheet
                IXLWorksheet tempSheet = worksheets.FirstOrDefault(ws => ws.Name == sheetBlockDataName);
                if (tempSheet == null)
                {
                    _logger.Warning($"{LogPrefix} Sheet '{sheetBlockDataName}' not found");
                    return false;
                }

                // Process all extended data blocks
                AddToExtendedDataFC(tempSheet);
                AddToExtendedDataSupportBD(tempSheet);
                AddToExtendedDataUserConstants(tempSheet);
                AddToExtendedDataFCValues(tempSheet);

                _logger.Information($"{LogPrefix} Finished reading extended data");
                return true;
            }
            catch (Exception ex)
            {
                _logger.Error($"{LogPrefix} Error in excel data processing: {ex.Message}");
                return false;
            }
        }

        // Private Methods
        /// <summary>
        /// Initializes or retrieves a PLC structure based on the PLC name from a row.
        /// </summary>
        /// <param name="row">The row number in the main worksheet.</param>
        /// <returns>The initialized or existing <see cref="dataPLC"/> instance.</returns>
        private dataPLC InitializePLC(int row)
        {
            string plcName = mainWorkSheet.Cell(row, fieldMain["PLCName"]).GetValue<string>(); // Get PLC name
            dataPLC plc = blocksStruct.FirstOrDefault(item => item.namePLC == plcName); // Find existing PLC

            // Create new PLC if not found
            if (plc == null)
            {
                plc = new dataPLC
                {
                    namePLC = plcName,
                    Equipment = new List<dataEq>(),
                    instanceDB = new List<dataBlock>(),
                    dataFC = new List<dataFunction>(),
                    dataSupportBD = new List<dataSupportBD>(),
                    userConstant = new List<userConstant>()
                };
                blocksStruct.Add(plc);
                _logger.Information($"{LogPrefix} Created structure for PLC: {plcName}");
            }
            return plc;
        }
        /// <summary>
        /// Retrieves the equipment worksheet based on the equipment type from a row.
        /// </summary>
        /// <param name="row">The row number in the main worksheet.</param>
        /// <param name="typeEq">Output parameter for the equipment type.</param>
        /// <param name="eqName">Output parameter for the equipment name.</param>
        /// <param name="sheet">Output parameter for the equipment worksheet.</param>
        /// <returns>True if the worksheet is found, false otherwise.</returns>
        private bool GetEquipmentSheet(int row, out string typeEq, out string eqName, out IXLWorksheet sheet)
        {
            string tempTypeEq = mainWorkSheet.Cell(row, fieldMain["EqType"]).GetValue<string>(); // Get equipment type
            sheet = worksheets.FirstOrDefault(ws => ws.Name.Equals(tempTypeEq, StringComparison.OrdinalIgnoreCase)); // Find worksheet

            // Handle case where worksheet is not found
            if (sheet == null)
            {
                _logger.Warning($"{LogPrefix} No sheet found for equipment type: {tempTypeEq}");
                typeEq = tempTypeEq;
                eqName = string.Empty;
                return false;
            }

            typeEq = tempTypeEq;
            eqName = sheet.Name;
            return true;
        }
        /// <summary>
        /// Processes equipment data for a PLC from its worksheet.
        /// </summary>
        /// <param name="sheet">The equipment worksheet.</param>
        /// <param name="tempItem">The PLC structure to populate.</param>
        /// <param name="typeEq">The equipment type.</param>
        /// <param name="eqName">The equipment name.</param>
        private void ProcessEquipmentData(IXLWorksheet sheet, dataPLC tempItem, string typeEq, string eqName)
        {
            // Skip if equipment type already exists
            if (tempItem.Equipment.Any(item => item.typeEq == typeEq))
                return;

            // Create new equipment structure
            dataEq tempEq = new dataEq
            {
                typeEq = typeEq,
                isExtended = sheet.Cell(typeIsExtendedCell.Row, typeIsExtendedCell.Col).GetValue<string>() == "Extended",
                FB = new List<dataLibrary>(),
                dataTag = new List<dataTag>(),
                dataParameter = new List<dataParameter>(),
                dataExtSupportBlock = new List<dataExtSupportBlock>(),
                dataConstant = new List<userConstant>(),
                dataSupportBD = new List<dataSupportBD>(),
                dataDataBlockValue = new List<dataDataBlockValue>()
            };
            tempItem.Equipment.Add(tempEq);
            _logger.Information($"{LogPrefix} Created structure for equipment type: {typeEq}");

            PopulateEquipmentData(sheet, tempEq); // Populate equipment data
        }
        /// <summary>
        /// Populates the equipment structure with data from its worksheet.
        /// </summary>
        /// <param name="sheet">The equipment worksheet.</param>
        /// <param name="tempEq">The equipment structure to populate.</param>
        private void PopulateEquipmentData(IXLWorksheet sheet, dataEq tempEq)
        {
            // Populate all equipment data sections
            AddToEquipmentFBLibraries(sheet, tempEq);
            AddToEquipmentDataTags(sheet, tempEq);
            AddToEquipmentDataParameters(sheet, tempEq);
            AddToEquipmentExtSupportBlocks(sheet, tempEq);
            AddToEquipmentConstants(sheet, tempEq);
            AddToEquipmentSupportBDs(sheet, tempEq);
            AddToEquipmentDataBlockValues(sheet, tempEq);
        }
        /// <summary>
        /// Adds function block libraries to the equipment structure.
        /// </summary>
        /// <param name="sheet">The equipment worksheet.</param>
        /// <param name="tempEq">The equipment structure to populate.</param>
        private void AddToEquipmentFBLibraries(IXLWorksheet sheet, dataEq tempEq)
        {
            ProcessBlockData(sheet, typeCoordinates, "FB", r => tempEq.FB.Add(new dataLibrary
            {
                name = sheet.Cell(r, fieldTypeFcColumns["name"]).GetValue<string>(),
                group = sheet.Cell(r, fieldTypeFcColumns["group"]).GetValue<string>(),
                path = sheet.Cell(r, fieldTypeFcColumns["path"]).GetValue<string>(),
                isType = sheet.Cell(r, fieldTypeFcColumns["isType"]).GetValue<string>().ToLower() == "yes"
            }), "Added Function Blocks (FB)");
        }
        /// <summary>
        /// Adds data tags to the equipment structure.
        /// </summary>
        /// <param name="sheet">The equipment worksheet.</param>
        /// <param name="tempEq">The equipment structure to populate.</param>
        private void AddToEquipmentDataTags(IXLWorksheet sheet, dataEq tempEq)
        {
            ProcessBlockData(sheet, typeCoordinates, "DataTag", r =>
            {
                string typeValue = sheet.Cell(r, fieldTypeTagColumns["type"]).GetValue<string>();
                string tAdress = string.Empty;
                // Assign default address based on type
                if (!string.IsNullOrEmpty(typeValue))
                {
                    switch (typeValue.ToLower())
                    {
                        case "bool": tAdress = "M100.0"; break;
                        case "word": tAdress = "MW1000"; break;
                        case "int": tAdress = "MW1000"; break;
                    }
                }
                var tag = new dataTag
                {
                    name = sheet.Cell(r, fieldTypeTagColumns["name"]).GetValue<string>(),
                    link = sheet.Cell(r, fieldTypeTagColumns["link"]).GetValue<string>(),
                    type = typeValue,
                    adress = tAdress,
                    table = string.IsNullOrEmpty(sheet.Cell(r, fieldTypeTagColumns["table"]).GetValue<string>()) ? "@Eq_IOTable" : sheet.Cell(r, fieldTypeTagColumns["table"]).GetValue<string>(),
                    comment = sheet.Cell(r, fieldTypeTagColumns["comment"]).GetValue<string>(),
                    variant = sheet.Cell(r, fieldTypeTagColumns["variant"]).GetValue<string>()?.Split(',')?.ToList() ?? new List<string>()
                };
                tempEq.dataTag.Add(tag);
            }, "Added Data Tags");
        }
        /// <summary>
        /// Adds data parameters to the equipment structure.
        /// </summary>
        /// <param name="sheet">The equipment worksheet.</param>
        /// <param name="tempEq">The equipment structure to populate.</param>
        private void AddToEquipmentDataParameters(IXLWorksheet sheet, dataEq tempEq)
        {
            ProcessBlockData(sheet, typeCoordinates, "DataParameter", r => tempEq.dataParameter.Add(new dataParameter
            {
                name = sheet.Cell(r, fieldTypeParameterColumns["name"]).GetValue<string>(),
                type = sheet.Cell(r, fieldTypeParameterColumns["type"]).GetValue<string>(),
                link = sheet.Cell(r, fieldTypeParameterColumns["link"]).GetValue<string>()
            }), "Added Data Parameters");
        }
        /// <summary>
        /// Adds extended support blocks to the equipment structure.
        /// </summary>
        /// <param name="sheet">The equipment worksheet.</param>
        /// <param name="tempEq">The equipment structure to populate.</param>
        private void AddToEquipmentExtSupportBlocks(IXLWorksheet sheet, dataEq tempEq)
        {
            ProcessBlockData(sheet, typeCoordinates, "ExtSupportBlock", r =>
            {
                string numberValue = sheet.Cell(r, fieldTypeExtSupportBlockColumns["number"]).GetValue<string>();
                int.TryParse(numberValue, out int number); // Parse number, defaults to 0 if invalid
                var block = new dataExtSupportBlock
                {
                    name = sheet.Cell(r, fieldTypeExtSupportBlockColumns["name"]).GetValue<string>(),
                    instanceOfName = sheet.Cell(r, fieldTypeExtSupportBlockColumns["instanceOfName"]).GetValue<string>(),
                    number = number,
                    variant = sheet.Cell(r, fieldTypeExtSupportBlockColumns["variant"]).GetValue<string>()?.Split(',')?.ToList() ?? new List<string>()
                };
                tempEq.dataExtSupportBlock.Add(block);
            }, "Added Extended Support Blocks");
        }
        /// <summary>
        /// Adds constants to the equipment structure.
        /// </summary>
        /// <param name="sheet">The equipment worksheet.</param>
        /// <param name="tempEq">The equipment structure to populate.</param>
        private void AddToEquipmentConstants(IXLWorksheet sheet, dataEq tempEq)
        {
            ProcessBlockData(sheet, typeCoordinates, "Constant", r => tempEq.dataConstant.Add(new userConstant
            {
                name = sheet.Cell(r, fieldTypeConstantColumns["name"]).GetValue<string>(),
                type = sheet.Cell(r, fieldTypeConstantColumns["type"]).GetValue<string>(),
                value = sheet.Cell(r, fieldTypeConstantColumns["value"]).GetValue<string>(),
                table = sheet.Cell(r, fieldTypeConstantColumns["table"]).GetValue<string>()
            }), "Added Constants");
        }
        /// <summary>
        /// Adds support BDs to the equipment structure.
        /// </summary>
        /// <param name="sheet">The equipment worksheet.</param>
        /// <param name="tempEq">The equipment structure to populate.</param>
        private void AddToEquipmentSupportBDs(IXLWorksheet sheet, dataEq tempEq)
        {
            ProcessBlockData(sheet, typeCoordinates, "SupportBD", r =>
            {
                string numberValue = sheet.Cell(r, fieldTypeSupportBDColumns["number"]).GetValue<string>();
                int.TryParse(numberValue, out int number); // Parse number, defaults to 0 if invalid
                tempEq.dataSupportBD.Add(new dataSupportBD
                {
                    name = sheet.Cell(r, fieldTypeSupportBDColumns["name"]).GetValue<string>(),
                    number = number,
                    group = sheet.Cell(r, fieldTypeSupportBDColumns["group"]).GetValue<string>(),
                    path = sheet.Cell(r, fieldTypeSupportBDColumns["path"]).GetValue<string>(),
                    isType = sheet.Cell(r, fieldTypeSupportBDColumns["isType"]).GetValue<string>().ToLower() == "yes",
                    isRetain = sheet.Cell(r, fieldTypeSupportBDColumns["isRetain"]).GetValue<string>().ToLower() == "yes",
                    isOptimazed = sheet.Cell(r, fieldTypeSupportBDColumns["isOptimazed"]).GetValue<string>().ToLower() == "yes"
                });
            }, "Added Support BDs");
        }
        /// <summary>
        /// Adds data block values to the equipment structure.
        /// </summary>
        /// <param name="sheet">The equipment worksheet.</param>
        /// <param name="tempEq">The equipment structure to populate.</param>
        private void AddToEquipmentDataBlockValues(IXLWorksheet sheet, dataEq tempEq)
        {
            ProcessBlockData(sheet, typeCoordinates, "DataBlockValue", r => tempEq.dataDataBlockValue.Add(new dataDataBlockValue
            {
                name = sheet.Cell(r, fieldTypeDataBlockValueColumns["name"]).GetValue<string>(),
                type = sheet.Cell(r, fieldTypeDataBlockValueColumns["type"]).GetValue<string>(),
                DB = sheet.Cell(r, fieldTypeDataBlockValueColumns["DB"]).GetValue<string>()
            }), "Added Data Block Values");
        }
        /// <summary>
        /// Adds full object data (instance) to the PLC structure.
        /// </summary>
        /// <param name="sheet">The equipment worksheet.</param>
        /// <param name="tempItem">The PLC structure to populate.</param>
        /// <param name="row">The row number in the main worksheet.</param>
        /// <returns>True if the instance is successfully added, false otherwise.</returns>
        private bool AddFullObjectData(IXLWorksheet sheet, dataPLC tempItem, int row)
        {
            string typeEq = mainWorkSheet.Cell(row, fieldMain["EqType"]).GetValue<string>();
            bool isExtended = tempItem.Equipment.First(item => item.typeEq == typeEq).isExtended;
            // Create instance with appropriate naming convention
            var instance = new dataBlock
            {
                name = isExtended ? $"FC_{mainWorkSheet.Cell(row, fieldMain["EqName"]).GetValue<string>()}" : $"iDB-{typeEq}|{mainWorkSheet.Cell(row, fieldMain["EqName"]).GetValue<string>()}",
                comment = mainWorkSheet.Cell(row, fieldMain["EqComments"]).GetValue<string>(),
                area = mainWorkSheet.Cell(row, fieldMain["EqArea"]).GetValue<string>(),
                instanceOfName = mainWorkSheet.Cell(row, fieldMain["InstanceOfName"]).GetValue<string>(),
                group = mainWorkSheet.Cell(row, fieldMain["GroupDB"]).GetValue<string>(),
                nameFC = mainWorkSheet.Cell(row, fieldMain["NameFC"]).GetValue<string>(),
                typeEq = typeEq,
                nameEq = mainWorkSheet.Cell(row, fieldMain["EqName"]).GetValue<string>(),
                variant = new List<string>(),
                excelData = new List<excelData>()
            };

            // Parse PLC number
            string numberValue = mainWorkSheet.Cell(row, fieldMain["PLCNumber"]).GetValue<string>();
            if (!int.TryParse(numberValue, out int number))
            {
                _logger.Error($"{LogPrefix} Failed to parse PLCNumber at row {row}, column {fieldMain["PLCNumber"]}: {numberValue}");
                return false;
            }
            instance.number = number;

            // Parse variant list
            string variantValue = mainWorkSheet.Cell(row, fieldMain["Variant"]).GetValue<string>();
            if (!string.IsNullOrEmpty(variantValue))
            {
                instance.variant.AddRange(variantValue.Replace('.', ',').Split(','));
            }

            // Populate excel data from main worksheet
            var ex = instance.excelData;
            foreach (var data in fieldMain)
            {
                ex.Add(new excelData
                {
                    name = data.Key,
                    column = data.Value,
                    value = mainWorkSheet.Cell(row, data.Value).GetValue<string>()
                });
            }

            tempItem.instanceDB.Add(instance);
            _logger.Information($"{LogPrefix} Created Instance for: {ex.FirstOrDefault(item => item.name == "EqName")?.value}");
            return true;
        }
        /// <summary>
        /// Adds function data to the PLC structure from the extended data worksheet.
        /// </summary>
        /// <param name="sheet">The extended data worksheet.</param>
        private void AddToExtendedDataFC(IXLWorksheet sheet)
        {
            ProcessBlockData(sheet, extendedCoordinates, "FCs", row =>
            {
                var PLC = blocksStruct.FirstOrDefault(plc => plc.namePLC == sheet.Cell(row, fieldExtendedBlockData["PLCName"]).GetValue<string>());
                if (PLC != null && !PLC.dataFC.Any(fc => fc.name == sheet.Cell(row, fieldExtendedBlockData["Block Name"]).GetValue<string>()))
                {
                    string numberValue = sheet.Cell(row, fieldExtendedBlockData["Block Number"]).GetValue<string>();
                    if (!int.TryParse(numberValue, out int number))
                    {
                        _logger.Error($"{LogPrefix} Failed to parse Block Number at row {row}: {numberValue}");
                        return;
                    }

                    PLC.dataFC.Add(new dataFunction
                    {
                        name = sheet.Cell(row, fieldExtendedBlockData["Block Name"]).GetValue<string>(),
                        number = number,
                        group = sheet.Cell(row, fieldExtendedBlockData["Block Group"]).GetValue<string>(),
                        code = new StringBuilder($"FUNCTION \"{sheet.Cell(row, fieldExtendedBlockData["Block Name"]).GetValue<string>()}\" : Void\r\n")
                            .AppendLine("{ S7_Optimized_Access := 'TRUE' }")
                            .AppendLine("AUTHOR : SE")
                            .AppendLine("FAMILY : Constructor"),
                        dataExtFCValue = new List<dataExtFCValue>()
                    });
                    _logger.Information($"{LogPrefix} Created template for function block: {sheet.Cell(row, fieldExtendedBlockData["Block Name"]).GetValue<string>()}");
                }
            }, "Added all support FCs");
        }
        /// <summary>
        /// Adds support block data to the PLC structure from the extended data worksheet.
        /// </summary>
        /// <param name="sheet">The extended data worksheet.</param>
        private void AddToExtendedDataSupportBD(IXLWorksheet sheet)
        {
            ProcessBlockData(sheet, extendedCoordinates, "SupportBDs", row =>
            {
                var PLC = blocksStruct.FirstOrDefault(plc => plc.namePLC == sheet.Cell(row, fieldExtendedSupportDB["PLCName"]).GetValue<string>());
                if (PLC != null && !PLC.dataSupportBD.Any(s => s.name == sheet.Cell(row, fieldExtendedSupportDB["Block Name"]).GetValue<string>()))
                {
                    PLC.dataSupportBD.Add(new dataSupportBD
                    {
                        name = sheet.Cell(row, fieldExtendedSupportDB["Block Name"]).GetValue<string>(),
                        group = sheet.Cell(row, fieldExtendedSupportDB["Block Group"]).GetValue<string>(),
                        type = sheet.Cell(row, fieldExtendedSupportDB["Block Type"]).GetValue<string>(),
                        path = sheet.Cell(row, fieldExtendedSupportDB["Path"]).GetValue<string>(),
                        isType = sheet.Cell(row, fieldExtendedSupportDB["Types"]).GetValue<string>().ToLower() == "yes",
                        isRetain = false
                    });
                    _logger.Information($"{LogPrefix} Created support BD in DB: {sheet.Cell(row, fieldExtendedSupportDB["Block Name"]).GetValue<string>()}");
                }
            }, "Added all support BD");
        }
        /// <summary>
        /// Adds user constants to the PLC structure from the extended data worksheet.
        /// </summary>
        /// <param name="sheet">The extended data worksheet.</param>
        private void AddToExtendedDataUserConstants(IXLWorksheet sheet)
        {
            ProcessBlockData(sheet, extendedCoordinates, "UserConstants", row =>
            {
                var PLC = blocksStruct.FirstOrDefault(plc => plc.namePLC == sheet.Cell(row, fieldExtendedUserConstant["PLCName"]).GetValue<string>());
                if (PLC != null && !PLC.userConstant.Any(s => s.name == sheet.Cell(row, fieldExtendedUserConstant["Block Name"]).GetValue<string>()))
                {
                    PLC.userConstant.Add(new userConstant
                    {
                        name = sheet.Cell(row, fieldExtendedUserConstant["Block Name"]).GetValue<string>(),
                        type = sheet.Cell(row, fieldExtendedUserConstant["Block Type"]).GetValue<string>(),
                        value = sheet.Cell(row, fieldExtendedUserConstant["Value"]).GetValue<string>()
                    });
                    _logger.Information($"{LogPrefix} Created user constant: {sheet.Cell(row, fieldExtendedUserConstant["Block Name"]).GetValue<string>()}");
                }
            }, "Added all support user constants");
        }
        /// <summary>
        /// Adds extended function values and builds function code in the PLC structure.
        /// </summary>
        /// <param name="sheet">The extended data worksheet.</param>
        private void AddToExtendedDataFCValues(IXLWorksheet sheet)
        {
            ProcessBlockData(sheet, extendedCoordinates, "ExtFCValues", row =>
            {
                var PLC = blocksStruct.FirstOrDefault(plc => plc.namePLC == sheet.Cell(row, fieldExtendedExtFCValue["PLCName"]).GetValue<string>());
                if (PLC != null)
                {
                    var fc = PLC.dataFC.FirstOrDefault(f => f.name == sheet.Cell(row, fieldExtendedExtFCValue["FC"]).GetValue<string>());
                    if (fc != null && !fc.dataExtFCValue.Any(value => value.name == sheet.Cell(row, fieldExtendedExtFCValue["Name"]).GetValue<string>()))
                    {
                        fc.dataExtFCValue.Add(new dataExtFCValue
                        {
                            name = sheet.Cell(row, fieldExtendedExtFCValue["Name"]).GetValue<string>(),
                            type = sheet.Cell(row, fieldExtendedExtFCValue["Type"]).GetValue<string>(),
                            IO = sheet.Cell(row, fieldExtendedExtFCValue["IO"]).GetValue<string>(),
                            Comments = sheet.Cell(row, fieldExtendedExtFCValue["Comments"]).GetValue<string>()
                        });
                    }
                }
            }, "Added all extended FC values");

            // Build code for each function in each PLC
            foreach (var PLC in BlocksStruct)
            {
                foreach (var FC in PLC.dataFC)
                {
                    AppendVariableSection(FC.code, FC.dataExtFCValue, "I", "VAR_INPUT");
                    AppendVariableSection(FC.code, FC.dataExtFCValue, "O", "VAR_OUTPUT");
                    AppendVariableSection(FC.code, FC.dataExtFCValue, "IO", "VAR_IN_OUT");
                    AppendVariableSection(FC.code, FC.dataExtFCValue, "T", "VAR_TEMP");
                    AppendVariableSection(FC.code, FC.dataExtFCValue, "C", "VAR_CONSTANT");

                    FC.code.AppendLine("BEGIN"); // Mark the start of the function body
                    _logger.Information($"{LogPrefix} Extended values for function: {FC.name}");
                }
            }
        }
        /// <summary>
        /// Processes a block of data from a worksheet and applies an action to each row.
        /// </summary>
        /// <param name="sheet">The worksheet to process.</param>
        /// <param name="coordinate">The coordinate dictionary for block locations.</param>
        /// <param name="blockType">The type of block to process (e.g., "FCs").</param>
        /// <param name="addItem">The action to perform on each row.</param>
        /// <param name="successMessage">The message to log on successful completion.</param>
        private void ProcessBlockData(IXLWorksheet sheet, Dictionary<string, (int Row, int StartCol)> coordinate, string blockType, Action<int> addItem, string successMessage)
        {
            // Get block range or exit if invalid
            if (!GetBlockRange(sheet, blockType, coordinate, out int numRowBlock, out int scopeRowBlock) || scopeRowBlock <= 0)
                return;

            // Process each row in the block
            for (int r = numRowBlock + 2; r < numRowBlock + 2 + scopeRowBlock; r++)
            {
                try
                {
                    addItem(r); // Apply the provided action to the row
                }
                catch (Exception ex)
                {
                    _logger.Error($"{LogPrefix} Error processing row {r} in block {blockType}: {ex.Message}");
                }
            }
            _logger.Debug($"{LogPrefix} {successMessage} for equipment type: {sheet.Name}");
        }
        /// <summary>
        /// Retrieves the range (start row and scope) of a block in a worksheet.
        /// </summary>
        /// <param name="sheet">The worksheet to query.</param>
        /// <param name="blockType">The type of block (e.g., "FCs").</param>
        /// <param name="coordinate">The coordinate dictionary for block locations.</param>
        /// <param name="numRowBlock">Output parameter for the starting row of the block.</param>
        /// <param name="scopeRowBlock">Output parameter for the number of rows in the block.</param>
        /// <returns>True if the range is successfully retrieved, false otherwise.</returns>
        private bool GetBlockRange(IXLWorksheet sheet, string blockType, Dictionary<string, (int Row, int StartCol)> coordinate, out int numRowBlock, out int scopeRowBlock)
        {
            numRowBlock = 0;
            scopeRowBlock = 0;
            var (row, startCol) = coordinate[blockType]; // Get block coordinates

            // Parse starting row number
            string numValue = sheet.Cell(row, startCol).GetValue<string>();
            if (!int.TryParse(numValue, out numRowBlock))
            {
                _logger.Warning($"{LogPrefix} Failed to parse numRowBlock at ({row}, {startCol}) for {blockType}: {numValue}");
                return false;
            }

            // Parse scope (number of rows)
            string scopeValue = sheet.Cell(row, startCol + 1).GetValue<string>();
            if (!int.TryParse(scopeValue, out scopeRowBlock))
            {
                _logger.Warning($"{LogPrefix} Failed to parse scopeRowBlock at ({row}, {startCol + 1}) for {blockType}: {scopeValue}");
                return false;
            }

            return true;
        }
        /// <summary>
        /// Appends a section of variable declarations to a function's code.
        /// </summary>
        /// <param name="code">The StringBuilder containing the function code.</param>
        /// <param name="values">The list of extended function values.</param>
        /// <param name="ioType">The IO type to filter by (e.g., "I" for inputs).</param>
        /// <param name="sectionHeader">The header for the section (e.g., "VAR_INPUT").</param>
        private void AppendVariableSection(StringBuilder code, List<dataExtFCValue> values, string ioType, string sectionHeader)
        {
            var filteredValues = values.Where(v => v.IO == ioType).ToList(); // Filter values by IO type
            if (filteredValues.Any())
            {
                code.AppendLine(sectionHeader); // Add section header
                foreach (var value in filteredValues)
                    code.AppendLine($"\"{value.name}\" : {value.type};   // {value.Comments}"); // Add variable declaration
                code.AppendLine("END_VAR"); // Close section
            }
        }
    }
}