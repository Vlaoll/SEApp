using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using ClosedXML.Excel;
using Microsoft.Win32;
using Serilog;
using seConfSW.Domain.Models;

namespace seConfSW
{
    public class ExcelDataReader
    {
        private string excelPath = string.Empty;
        private string message = string.Empty;
        private XLWorkbook workbook;
        private IXLWorksheet mainWorkSheet;       
        private List<IXLWorksheet> worksheets;
        private readonly ILogger _logger;
        private const string LogPrefix = "[Excel]";

        private readonly List<dataPLC> blocksStruct = new List<dataPLC>();
        public string Message => message;
        public List<dataPLC> BlocksStruct => blocksStruct;

        private readonly Dictionary<string, int> fields = new Dictionary<string, int>
        {
            {"Status", 1}, {"PLCName", 2}, {"PrjNum", 3}, {"EqName", 4}, {"EqArea", 5}, {"EqComments", 6},
            {"EqType", 7}, {"PLCNumber", 8}, {"InstanceOfName", 9}, {"GroupDB", 10}, {"NameFC", 11},
            {"Variant", 12}, {"ObjName", 13}, {"PicNum", 14}, {"ObjTagName", 15}, {"HMIType", 16},
            {"TypicalPDL", 17}, {"WorkPDL", 18}, {"X", 19}, {"Y", 20}, {"Width", 21}, {"Height", 22},
            {"ScaleMsg", 23}, {"Addition_01", 24}, {"Addition_02", 25}, {"Addition_03", 26},
            {"Addition_04", 27}, {"Addition_05", 28}, {"Addition_06", 29}, {"Addition_07", 30},
            {"Addition_08", 31}, {"Addition_09", 32}, {"Addition_10", 33}
        };

        private readonly Dictionary<string, int> fieldsBlockData = new Dictionary<string, int>
        {
            {"Block Name", 2}, {"Block Number", 3}, {"Block Group", 4}, {"PLCName", 5}, {"Block Type", 6}
        };

        private readonly Dictionary<string, int> fieldsSupportDB = new Dictionary<string, int>
        {
            {"Block Name", 2}, {"Block Group", 3}, {"PLCName", 4}, {"Block Type", 5}, {"Path", 6}, {"Types", 7}
        };

        private readonly Dictionary<string, int> fieldsUserConstant = new Dictionary<string, int>
        {
            {"Block Name", 2}, {"Block Type", 3}, {"Value", 4}, {"PLCName", 5}
        };

        private readonly Dictionary<string, int> fieldsExtFCValue = new Dictionary<string, int>
        {
            {"Name", 2}, {"Type", 3}, {"IO", 4}, {"Comments", 5}, {"FC", 6}, {"PLCName", 7}
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
        private readonly Dictionary<string, int> fbColumns = new Dictionary<string, int>
        {
            {"name", 2}, {"group", 3}, {"path", 4}, {"isType", 5}
        };

        private readonly Dictionary<string, int> dataTagColumns = new Dictionary<string, int>
        {
            {"name", 2}, {"link", 3}, {"type", 4}, {"table", 5}, {"comment", 6}, {"variant", 7}
        };

        private readonly Dictionary<string, int> dataParameterColumns = new Dictionary<string, int>
        {
            {"name", 2}, {"type", 3}, {"link", 4}
        };

        private readonly Dictionary<string, int> extSupportBlockColumns = new Dictionary<string, int>
        {
            {"name", 2}, {"instanceOfName", 3}, {"number", 4}, {"variant", 5}
        };

        private readonly Dictionary<string, int> constantColumns = new Dictionary<string, int>
        {
            {"name", 2}, {"type", 3}, {"value", 4}, {"table", 5}
        };

        private readonly Dictionary<string, int> supportBDColumns = new Dictionary<string, int>
        {
            {"name", 2}, {"number", 3}, {"group", 4}, {"path", 5}, {"isType", 6}, {"isRetain", 7}, {"isOptimazed", 8}
        };

        private readonly Dictionary<string, int> dataBlockValueColumns = new Dictionary<string, int>
        {
            {"name", 2}, {"type", 3}, {"DB", 4}
        };


        private readonly (int Row, int Col) IsExtendedCell = (2, 5);

        public ExcelDataReader(ILogger logger = null)
        {
            _logger = logger ?? Log.ForContext<ExcelDataReader>();
            worksheets = new List<IXLWorksheet>();
            _logger.Information("ExcelDataReader initialized successfully.");
        }

        public string SearchProject(string filter = "Excel |*.xlsx;*.xlsm")
        {
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

            if (!string.IsNullOrEmpty(projectPath))
            {
                message = $"{LogPrefix} Opening excel file {projectPath}";
                _logger.Information(message);
            }
            return projectPath;
        }

        public bool OpenExcelFile(string filename, string mainSheetName = "Main")
        {
            try
            {
                excelPath = new FileInfo(filename).FullName;
                workbook = new XLWorkbook(excelPath);
                worksheets = workbook.Worksheets.ToList();
                mainWorkSheet = worksheets.FirstOrDefault(ws => ws.Name == mainSheetName);

                if (mainWorkSheet == null)
                {
                    message = $"{LogPrefix} Main sheet '{mainSheetName}' not found in {excelPath}";
                    _logger.Warning(message);
                    return false;
                }

                message = $"{LogPrefix} Master Excel {excelPath} is opened";
                _logger.Information(message);
                return true;
            }
            catch (Exception ex)
            {
                message = $"{LogPrefix} Error while opening excel: {ex.Message}";
                _logger.Error(ex, message);
                return false;
            }
        }

        public bool CloseExcelFile(bool save = false)
        {
            if (workbook == null)
            {
                message = $"{LogPrefix} No open workbook to close";
                _logger.Warning(message);
                return false;
            }

            try
            {
                if (save) workbook.Save();
                workbook.Dispose();
                workbook = null;
                message = $"{LogPrefix} Master Excel {excelPath} is closed";
                _logger.Information(message);
                return true;
            }
            catch (Exception ex)
            {
                message = $"{LogPrefix} Error while closing excel: {ex.Message}";
                _logger.Error(ex, message);
                return false;
            }
        }

        public bool ReadExcelObjectData(string status, int maxInstanceCount = 250)
        {
            if (mainWorkSheet == null)
            {
                message = $"{LogPrefix} Main worksheet is not initialized";
                _logger.Error(message);
                return false;
            }

            try
            {
                var last = mainWorkSheet.LastCellUsed();
                bool isAnySuccess = false;

                for (int row = 3; row <= last.Address.RowNumber; row++)
                {
                    if (!mainWorkSheet.Cell(row, fields["Status"]).GetValue<string>().Contains(status))
                        continue;

                    try
                    {
                        dataPLC tempItem = InitializePLC(row);
                        if (maxInstanceCount > 0 && tempItem.instanceDB.Count >= maxInstanceCount)
                            continue;

                        if (!GetEquipmentSheet(row, out string typeEq, out string eqName,out IXLWorksheet eqWorkSheet))
                            continue;

                        ProcessEquipmentData(eqWorkSheet,tempItem, typeEq, eqName);
                        if (AddFullObjectData(eqWorkSheet, tempItem, row))
                            isAnySuccess = true;
                    }
                    catch (Exception ex)
                    {
                        message = $"{LogPrefix} Error processing row {row}: {ex.Message}";
                        _logger.Error(ex, message);
                    }
                }
                return isAnySuccess;
            }
            catch (Exception ex)
            {
                message = $"{LogPrefix} Critical error in excel data processing: {ex.Message}";
                _logger.Error(ex, message);
                return false;
            }
        }

        private dataPLC InitializePLC(int row)
        {
            string plcName = mainWorkSheet.Cell(row, fields["PLCName"]).GetValue<string>();
            dataPLC plc = blocksStruct.FirstOrDefault(item => item.namePLC == plcName);

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
                message = $"{LogPrefix} Created structure for PLC: {plcName}";
                _logger.Information(message);
            }
            return plc;
        }

        private bool GetEquipmentSheet(int row, out string typeEq, out string eqName, out IXLWorksheet sheet)
        {
            string tempTypeEq = mainWorkSheet.Cell(row, fields["EqType"]).GetValue<string>();
            sheet = worksheets.FirstOrDefault(ws => ws.Name.Equals(tempTypeEq, StringComparison.OrdinalIgnoreCase));

            if (sheet == null)
            {
                message = $"{LogPrefix} No sheet found for equipment type: {tempTypeEq}";
                _logger.Warning(message);
                typeEq = tempTypeEq;
                eqName = string.Empty;
                return false;
            }

            typeEq = tempTypeEq;
            eqName = sheet.Name;
            return true;
        }

        private void ProcessEquipmentData(IXLWorksheet sheet,dataPLC tempItem, string typeEq, string eqName)
        {
            if (tempItem.Equipment.Any(item => item.typeEq == typeEq))
                return;

            dataEq tempEq = new dataEq
            {
                typeEq = typeEq,
                isExtended = sheet.Cell(IsExtendedCell.Row, IsExtendedCell.Col).GetValue<string>() == "Extended",
                FB = new List<dataLibrary>(),
                dataTag = new List<dataTag>(),
                dataParameter = new List<dataParameter>(),
                dataExtSupportBlock = new List<dataExtSupportBlock>(),
                dataConstant = new List<userConstant>(),
                dataSupportBD = new List<dataSupportBD>(),
                dataDataBlockValue = new List<dataDataBlockValue>()
            };
            tempItem.Equipment.Add(tempEq);
            message = $"{LogPrefix} Created structure for equipment type: {typeEq}";
            _logger.Information(message);

            PopulateEquipmentData(sheet,tempEq);
        }

        private void PopulateEquipmentData(IXLWorksheet sheet,dataEq tempEq)
        {
            AddToEquipmentFBLibraries(sheet, tempEq);
            AddToEquipmentDataTags(sheet, tempEq);
            AddToEquipmentDataParameters(sheet, tempEq);
            AddToEquipmentExtSupportBlocks(sheet, tempEq);
            AddToEquipmentConstants(sheet, tempEq);
            AddToEquipmentSupportBDs(sheet, tempEq);
            AddToEquipmentDataBlockValues(sheet, tempEq);
        }

        private void AddToEquipmentFBLibraries(IXLWorksheet sheet,dataEq tempEq)
        {
            ProcessBlockData(sheet,typeCoordinates, "FB", r => tempEq.FB.Add(new dataLibrary
            {
                name = sheet.Cell(r, fbColumns["name"]).GetValue<string>(),
                group = sheet.Cell(r, fbColumns["group"]).GetValue<string>(),
                path =  sheet.Cell(r, fbColumns["path"]).GetValue<string>(),
                isType = sheet.Cell(r, fbColumns["isType"]).GetValue<string>().ToLower() == "yes"
            }), "Added Function Blocks (FB)");
        }

        private void AddToEquipmentDataTags(IXLWorksheet sheet,dataEq tempEq)
        {
            ProcessBlockData(sheet, typeCoordinates, "DataTag",  r =>
            {
                string typeValue = sheet.Cell(r, dataTagColumns["type"]).GetValue<string>();
                string tAdress = string.Empty;
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
                    name = sheet.Cell(r, dataTagColumns["name"]).GetValue<string>(),
                    link = sheet.Cell(r, dataTagColumns["link"]).GetValue<string>(),
                    type = typeValue,
                    adress = tAdress,
                    table = string.IsNullOrEmpty(sheet.Cell(r, dataTagColumns["table"]).GetValue<string>()) ? "@Eq_IOTable" : sheet.Cell(r, dataTagColumns["table"]).GetValue<string>(),
                    comment = sheet.Cell(r, dataTagColumns["comment"]).GetValue<string>(),
                    variant = sheet.Cell(r, dataTagColumns["variant"]).GetValue<string>()?.Split(',')?.ToList() ?? new List<string>()
                };
                tempEq.dataTag.Add(tag);
            }, "Added Data Tags");
        }

        private void AddToEquipmentDataParameters(IXLWorksheet sheet,dataEq tempEq)
        {
            ProcessBlockData(sheet, typeCoordinates, "DataParameter", r => tempEq.dataParameter.Add(new dataParameter
            {
                name = sheet.Cell(r, dataParameterColumns["name"]).GetValue<string>(),
                type = sheet.Cell(r, dataParameterColumns["type"]).GetValue<string>(),
                link = sheet.Cell(r, dataParameterColumns["link"]).GetValue<string>()
            }), "Added Data Parameters");
        }

        private void AddToEquipmentExtSupportBlocks(IXLWorksheet sheet,dataEq tempEq)
        {
            ProcessBlockData(sheet, typeCoordinates, "ExtSupportBlock", r =>
            {
                string numberValue = sheet.Cell(r, extSupportBlockColumns["number"]).GetValue<string>();
                int.TryParse(numberValue, out int number);
                var block = new dataExtSupportBlock
                {
                    name = sheet.Cell(r, extSupportBlockColumns["name"]).GetValue<string>(),
                    instanceOfName = sheet.Cell(r, extSupportBlockColumns["instanceOfName"]).GetValue<string>(),
                    number = number,
                    variant = sheet.Cell(r, extSupportBlockColumns["variant"]).GetValue<string>()?.Split(',')?.ToList() ?? new List<string>()
                };
                tempEq.dataExtSupportBlock.Add(block);
            }, "Added Extended Support Blocks");
        }

        private void AddToEquipmentConstants(IXLWorksheet sheet,dataEq tempEq)
        {
            ProcessBlockData(sheet, typeCoordinates, "Constant", r => tempEq.dataConstant.Add(new userConstant
            {
                name = sheet.Cell(r, constantColumns["name"]).GetValue<string>(),
                type = sheet.Cell(r, constantColumns["type"]).GetValue<string>(),
                value = sheet.Cell(r, constantColumns["value"]).GetValue<string>(),
                table = sheet.Cell(r, constantColumns["table"]).GetValue<string>()
            }), "Added Constants");
        }

        private void AddToEquipmentSupportBDs(IXLWorksheet sheet,dataEq tempEq)
        {
            ProcessBlockData(sheet, typeCoordinates, "SupportBD", r =>
            {
                string numberValue = sheet.Cell(r, supportBDColumns["number"]).GetValue<string>();
                int.TryParse(numberValue, out int number);
                tempEq.dataSupportBD.Add(new dataSupportBD
                {
                    name = sheet.Cell(r, supportBDColumns["name"]).GetValue<string>(),
                    number = number,
                    group = sheet.Cell(r, supportBDColumns["group"]).GetValue<string>(),
                    path = sheet.Cell(r, supportBDColumns["path"]).GetValue<string>(),
                    isType = sheet.Cell(r, supportBDColumns["isType"]).GetValue<string>().ToLower() == "yes",
                    isRetain = sheet.Cell(r, supportBDColumns["isRetain"]).GetValue<string>().ToLower() == "yes",
                    isOptimazed = sheet.Cell(r, supportBDColumns["isOptimazed"]).GetValue<string>().ToLower() == "yes"
                });
            }, "Added Support BDs");
        }

        private void AddToEquipmentDataBlockValues(IXLWorksheet sheet,dataEq tempEq)
        {
            ProcessBlockData(sheet, typeCoordinates, "DataBlockValue",  r => tempEq.dataDataBlockValue.Add(new dataDataBlockValue
            {
                name = sheet.Cell(r, dataBlockValueColumns["name"]).GetValue<string>(),
                type = sheet.Cell(r, dataBlockValueColumns["type"]).GetValue<string>(),
                DB = sheet.Cell(r, dataBlockValueColumns["DB"]).GetValue<string>()
            }), "Added Data Block Values");
        }

        private bool AddFullObjectData(IXLWorksheet sheet,dataPLC tempItem, int row)
        {
            string typeEq = mainWorkSheet.Cell(row, fields["EqType"]).GetValue<string>();
            bool isExtended = tempItem.Equipment.First(item => item.typeEq == typeEq).isExtended;
            var instance = new dataBlock
            {
                name = isExtended ? $"FC_{mainWorkSheet.Cell(row, fields["EqName"]).GetValue<string>()}" : $"iDB-{typeEq}|{mainWorkSheet.Cell(row, fields["EqName"]).GetValue<string>()}",
                comment = mainWorkSheet.Cell(row, fields["EqComments"]).GetValue<string>(),
                area = mainWorkSheet.Cell(row, fields["EqArea"]).GetValue<string>(),
                instanceOfName = mainWorkSheet.Cell(row, fields["InstanceOfName"]).GetValue<string>(),
                group = mainWorkSheet.Cell(row, fields["GroupDB"]).GetValue<string>(),
                nameFC = mainWorkSheet.Cell(row, fields["NameFC"]).GetValue<string>(),
                typeEq = typeEq,
                nameEq = mainWorkSheet.Cell(row, fields["EqName"]).GetValue<string>(),
                variant = new List<string>(),
                excelData = new List<excelData>()
            };

            string numberValue = mainWorkSheet.Cell(row, fields["PLCNumber"]).GetValue<string>();
            if (!int.TryParse(numberValue, out int number))
            {
                message = $"{LogPrefix} Failed to parse PLCNumber at row {row}, column {fields["PLCNumber"]}: {numberValue}";
                _logger.Warning(message);
                return false;
            }
            instance.number = number;

            string variantValue = mainWorkSheet.Cell(row, fields["Variant"]).GetValue<string>();
            if (!string.IsNullOrEmpty(variantValue))
            {
                instance.variant.AddRange(variantValue.Replace('.', ',').Split(','));
            }

            var ex = instance.excelData;
            foreach (var data in fields)
            {
                ex.Add(new excelData
                {
                    name = data.Key,
                    column = data.Value,
                    value = mainWorkSheet.Cell(row, data.Value).GetValue<string>()
                });
            }

            tempItem.instanceDB.Add(instance);
            message = $"{LogPrefix} Created Instance for: {ex.FirstOrDefault(item => item.name == "EqName")?.value}";
            _logger.Information(message);
            return true;
        }

       
        public bool ReadExcelExtendedData(string sheetBlockDataName = "PLCData")
        {
            try
            {
                IXLWorksheet tempSheet = worksheets.FirstOrDefault(ws => ws.Name == sheetBlockDataName);
                if (tempSheet == null)
                {
                    message = $"{LogPrefix} Sheet '{sheetBlockDataName}' not found";
                    _logger.Warning(message);
                    return false;
                }

                AddToExtendedDataFC(tempSheet);
                AddToExtendedDataSupportBD(tempSheet);
                AddToExtendedDataUserConstants(tempSheet);
                AddToExtendedDataFCValues(tempSheet);

                return true;
            }
            catch (Exception ex)
            {
                message = $"{LogPrefix} Error in excel data processing: {ex.Message}";
                _logger.Error(ex, message);
                return false;
            }
        }

        private void AddToExtendedDataFC(IXLWorksheet sheet)
        {
            ProcessBlockData(sheet, extendedCoordinates, "FCs", row =>
            {
                var PLC = blocksStruct.FirstOrDefault(plc => plc.namePLC == sheet.Cell(row, fieldsBlockData["PLCName"]).GetValue<string>());
                if (PLC != null && !PLC.dataFC.Any(fc => fc.name == sheet.Cell(row, fieldsBlockData["Block Name"]).GetValue<string>()))
                {
                    string numberValue = sheet.Cell(row, fieldsBlockData["Block Number"]).GetValue<string>();
                    if (!int.TryParse(numberValue, out int number))
                    {
                        message = $"{LogPrefix} Failed to parse Block Number at row {row}: {numberValue}";
                        _logger.Warning(message);
                        return;
                    }

                    PLC.dataFC.Add(new dataFunction
                    {
                        name = sheet.Cell(row, fieldsBlockData["Block Name"]).GetValue<string>(),
                        number = number,
                        group = sheet.Cell(row, fieldsBlockData["Block Group"]).GetValue<string>(),
                        code = new StringBuilder($"FUNCTION \"{sheet.Cell(row, fieldsBlockData["Block Name"]).GetValue<string>()}\" : Void\r\n")
                            .AppendLine("{ S7_Optimized_Access := 'TRUE' }")
                            .AppendLine("AUTHOR : SE")
                            .AppendLine("FAMILY : Constructor"),
                        dataExtFCValue = new List<dataExtFCValue>()
                    });
                    message = $"{LogPrefix} Created template for function block: {sheet.Cell(row, fieldsBlockData["Block Name"]).GetValue<string>()}";
                    _logger.Information(message);
                }
            }, "Added all support FCs");
        }

        private void AddToExtendedDataSupportBD(IXLWorksheet sheet)
        {
            ProcessBlockData(sheet, extendedCoordinates, "SupportBDs", row =>
            {
                var PLC = blocksStruct.FirstOrDefault(plc => plc.namePLC == sheet.Cell(row, fieldsSupportDB["PLCName"]).GetValue<string>());
                if (PLC != null && !PLC.dataSupportBD.Any(s => s.name == sheet.Cell(row, fieldsSupportDB["Block Name"]).GetValue<string>()))
                {
                    PLC.dataSupportBD.Add(new dataSupportBD
                    {
                        name = sheet.Cell(row, fieldsSupportDB["Block Name"]).GetValue<string>(),
                        group = sheet.Cell(row, fieldsSupportDB["Block Group"]).GetValue<string>(),
                        type = sheet.Cell(row, fieldsSupportDB["Block Type"]).GetValue<string>(),
                        path = sheet.Cell(row, fieldsSupportDB["Path"]).GetValue<string>(),
                        isType = sheet.Cell(row, fieldsSupportDB["Types"]).GetValue<string>().ToLower() == "yes",
                        isRetain = false
                    });
                    message = $"{LogPrefix} Created support BD in DB: {sheet.Cell(row, fieldsSupportDB["Block Name"]).GetValue<string>()}";
                    _logger.Information(message);
                }
            }, "Added all support BD");
        }

        private void AddToExtendedDataUserConstants(IXLWorksheet sheet)
        {
            ProcessBlockData(sheet, extendedCoordinates, "UserConstants", row =>
            {
                var PLC = blocksStruct.FirstOrDefault(plc => plc.namePLC == sheet.Cell(row, fieldsUserConstant["PLCName"]).GetValue<string>());
                if (PLC != null && !PLC.userConstant.Any(s => s.name == sheet.Cell(row, fieldsUserConstant["Block Name"]).GetValue<string>()))
                {
                    PLC.userConstant.Add(new userConstant
                    {
                        name = sheet.Cell(row, fieldsUserConstant["Block Name"]).GetValue<string>(),
                        type = sheet.Cell(row, fieldsUserConstant["Block Type"]).GetValue<string>(),
                        value = sheet.Cell(row, fieldsUserConstant["Value"]).GetValue<string>()
                    });
                    message = $"{LogPrefix} Created user constant: {sheet.Cell(row, fieldsUserConstant["Block Name"]).GetValue<string>()}";
                    _logger.Information(message);
                }
            }, "Added all support user constants");
        }

        private void AddToExtendedDataFCValues(IXLWorksheet sheet)
        {
            if (!GetBlockRange(sheet, "ExtFCValues",extendedCoordinates, out int numRowBlock, out int scopeRowBlock) || scopeRowBlock <= 0)
                return;

            for (int row = numRowBlock + 2; row < numRowBlock + 2 + scopeRowBlock; row++)
            {
                try
                {
                    var PLC = blocksStruct.FirstOrDefault(plc => plc.namePLC == sheet.Cell(row, fieldsExtFCValue["PLCName"]).GetValue<string>());
                    if (PLC != null)
                    {
                        var fc = PLC.dataFC.FirstOrDefault(f => f.name == sheet.Cell(row, fieldsExtFCValue["FC"]).GetValue<string>());
                        if (fc != null && !fc.dataExtFCValue.Any(value => value.name == sheet.Cell(row, fieldsExtFCValue["Name"]).GetValue<string>()))
                        {
                            fc.dataExtFCValue.Add(new dataExtFCValue
                            {
                                name = sheet.Cell(row, fieldsExtFCValue["Name"]).GetValue<string>(),
                                type = sheet.Cell(row, fieldsExtFCValue["Type"]).GetValue<string>(),
                                IO = sheet.Cell(row, fieldsExtFCValue["IO"]).GetValue<string>(),
                                Comments = sheet.Cell(row, fieldsExtFCValue["Comments"]).GetValue<string>()
                            });
                        }
                    }
                }
                catch (Exception ex)
                {
                    message = $"{LogPrefix} Error processing row {row} in ExtFCValues: {ex.Message}";
                    _logger.Warning(ex, message);
                }
            }

            foreach (var PLC in BlocksStruct)
            {
                foreach (var FC in PLC.dataFC)
                {
                    var inputs = FC.dataExtFCValue.Where(type => type.IO == "I").ToList();
                    if (inputs.Any())
                    {
                        FC.code.AppendLine("VAR_INPUT");
                        foreach (var value in inputs)
                            FC.code.AppendLine($"\"{value.name}\" : {value.type};   // {value.Comments}");
                        FC.code.AppendLine("END_VAR");
                    }

                    var outputs = FC.dataExtFCValue.Where(type => type.IO == "O").ToList();
                    if (outputs.Any())
                    {
                        FC.code.AppendLine("VAR_OUTPUT");
                        foreach (var value in outputs)
                            FC.code.AppendLine($"\"{value.name}\" : {value.type};   // {value.Comments}");
                        FC.code.AppendLine("END_VAR");
                    }

                    var inOuts = FC.dataExtFCValue.Where(type => type.IO == "IO").ToList();
                    if (inOuts.Any())
                    {
                        FC.code.AppendLine("VAR_IN_OUT");
                        foreach (var value in inOuts)
                            FC.code.AppendLine($"\"{value.name}\" : {value.type};   // {value.Comments}");
                        FC.code.AppendLine("END_VAR");
                    }

                    var temps = FC.dataExtFCValue.Where(type => type.IO == "T").ToList();
                    if (temps.Any())
                    {
                        FC.code.AppendLine("VAR_TEMP");
                        foreach (var value in temps)
                            FC.code.AppendLine($"\"{value.name}\" : {value.type};   // {value.Comments}");
                        FC.code.AppendLine("END_VAR");
                    }

                    var constants = FC.dataExtFCValue.Where(type => type.IO == "C").ToList();
                    if (constants.Any())
                    {
                        FC.code.AppendLine("VAR CONSTANT");
                        foreach (var value in constants)
                            FC.code.AppendLine($"\"{value.name}\" : {value.type};   // {value.Comments}");
                        FC.code.AppendLine("END_VAR");
                    }

                    FC.code.AppendLine("BEGIN");
                    message = $"{LogPrefix} Extended values for function: {FC.name}";
                    _logger.Information(message);
                }
            }
        }


        private void ProcessBlockData(IXLWorksheet sheet, Dictionary<string, (int Row, int StartCol)> coordinate, string blockType, Action<int> addItem, string successMessage)
        {
            if (!GetBlockRange(sheet, blockType, coordinate, out int numRowBlock, out int scopeRowBlock) || scopeRowBlock <= 0)
                return;

            for (int r = numRowBlock + 2; r < numRowBlock + 2 + scopeRowBlock; r++)
            {
                try
                {
                    addItem(r);
                }
                catch (Exception ex)
                {
                    message = $"{LogPrefix} Error processing row {r} in block {blockType}: {ex.Message}";
                    _logger.Warning(ex, message);
                }
            }
            message = $"{LogPrefix} {successMessage} for equipment type: {sheet.Name}";
            _logger.Information(message);
        }
        private bool GetBlockRange(IXLWorksheet sheet, string blockType, Dictionary<string, (int Row, int StartCol)> coordinate, out int numRowBlock, out int scopeRowBlock)
        {
            numRowBlock = 0;
            scopeRowBlock = 0;
            var (row, startCol) = coordinate[blockType];

            string numValue = sheet.Cell(row, startCol).GetValue<string>();
            if (!int.TryParse(numValue, out numRowBlock))
            {
                message = $"{LogPrefix} Failed to parse numRowBlock at ({row}, {startCol}) for {blockType}: {numValue}";
                _logger.Warning(message);
                return false;
            }

            string scopeValue = sheet.Cell(row, startCol + 1).GetValue<string>();
            if (!int.TryParse(scopeValue, out scopeRowBlock))
            {
                message = $"{LogPrefix} Failed to parse scopeRowBlock at ({row}, {startCol + 1}) for {blockType}: {scopeValue}";
                _logger.Warning(message);
                return false;
            }

            return true;
        }

    }
}