using ClosedXML.Excel;
using Microsoft.Office.Interop.Excel;
using seConfSW.Domain.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TIAApp.Application.Services
{
    class Class1
    {
        public bool ReadExcelExtendedData(string sheetBlockDataNane = "PLCData")
        {
            try
            {
                IXLWorksheet tempSheet = worksheets.FirstOrDefault(ws => ws.Name == sheetBlockDataNane);
                if (tempSheet == null)
                {
                    message = $"{LogPrefix} Sheet '{sheetBlockDataNane}' not found";
                    _logger.Warning(message);
                    return false;
                }

                string numRowBlockValue = tempSheet.Cell(3, 3).GetValue<string>();
                string scopeRowBlockValue = tempSheet.Cell(3, 4).GetValue<string>();
                if (int.TryParse(numRowBlockValue, out int numRowBlock) && int.TryParse(scopeRowBlockValue, out int scopeRowBlock) && scopeRowBlock > 0)
                {
                    for (int row = numRowBlock + 2; row < numRowBlock + 2 + scopeRowBlock; row++)
                    {
                        var PLC = blocksStruct.FirstOrDefault(plc => plc.namePLC == tempSheet.Cell(row, fieldsBlockData["PLCName"]).GetValue<string>());
                        if (PLC != null && !PLC.dataFC.Any(fc => fc.name == tempSheet.Cell(row, fieldsBlockData["Block Name"]).GetValue<string>()))
                        {
                            string numberValue = tempSheet.Cell(row, fieldsBlockData["Block Number"]).GetValue<string>();
                            if (!int.TryParse(numberValue, out int number))
                            {
                                message = $"{LogPrefix} Failed to parse Block Number at row {row}: {numberValue}";
                                _logger.Warning(message);
                                continue;
                            }

                            PLC.dataFC.Add(new dataFunction
                            {
                                name = tempSheet.Cell(row, fieldsBlockData["Block Name"]).GetValue<string>(),
                                number = number,
                                group = tempSheet.Cell(row, fieldsBlockData["Block Group"]).GetValue<string>(),
                                code = new StringBuilder($"FUNCTION \"{tempSheet.Cell(row, fieldsBlockData["Block Name"]).GetValue<string>()}\" : Void\r\n")
                                    .AppendLine("{ S7_Optimized_Access := 'TRUE' }")
                                    .AppendLine("AUTHOR : SE")
                                    .AppendLine("FAMILY : Constructor"),
                                dataExtFCValue = new List<dataExtFCValue>()
                            });
                            message = $"{LogPrefix} Created template for function block: {tempSheet.Cell(row, fieldsBlockData["Block Name"]).GetValue<string>()}";
                            _logger.Information(message);
                        }
                    }
                    message = $"{LogPrefix} Added all support FCs";
                    _logger.Information(message);
                }

                numRowBlockValue = tempSheet.Cell(4, 3).GetValue<string>();
                scopeRowBlockValue = tempSheet.Cell(4, 4).GetValue<string>();
                if (int.TryParse(numRowBlockValue, out numRowBlock) && int.TryParse(scopeRowBlockValue, out scopeRowBlock) && scopeRowBlock > 0)
                {
                    for (int row = numRowBlock + 2; row < numRowBlock + 2 + scopeRowBlock; row++)
                    {
                        var PLC = blocksStruct.FirstOrDefault(plc => plc.namePLC == tempSheet.Cell(row, fieldsSupportDB["PLCName"]).GetValue<string>());
                        if (PLC != null && !PLC.dataSupportBD.Any(s => s.name == tempSheet.Cell(row, fieldsSupportDB["Block Name"]).GetValue<string>()))
                        {
                            PLC.dataSupportBD.Add(new dataSupportBD
                            {
                                name = tempSheet.Cell(row, fieldsSupportDB["Block Name"]).GetValue<string>(),
                                group = tempSheet.Cell(row, fieldsSupportDB["Block Group"]).GetValue<string>(),
                                type = tempSheet.Cell(row, fieldsSupportDB["Block Type"]).GetValue<string>(),
                                path = tempSheet.Cell(row, fieldsSupportDB["Path"]).GetValue<string>(),
                                isType = tempSheet.Cell(row, fieldsSupportDB["Types"]).GetValue<string>().ToLower() == "yes",
                                isRetain = false
                            });
                            message = $"{LogPrefix} Created support BD in DB: {tempSheet.Cell(row, fieldsSupportDB["Block Name"]).GetValue<string>()}";
                            _logger.Information(message);
                        }
                    }
                    message = $"{LogPrefix} Added all support BD";
                    _logger.Information(message);
                }

                numRowBlockValue = tempSheet.Cell(5, 3).GetValue<string>();
                scopeRowBlockValue = tempSheet.Cell(5, 4).GetValue<string>();
                if (int.TryParse(numRowBlockValue, out numRowBlock) && int.TryParse(scopeRowBlockValue, out scopeRowBlock) && scopeRowBlock > 0)
                {
                    for (int row = numRowBlock + 2; row < numRowBlock + 2 + scopeRowBlock; row++)
                    {
                        var PLC = blocksStruct.FirstOrDefault(plc => plc.namePLC == tempSheet.Cell(row, fieldsUserConstant["PLCName"]).GetValue<string>());
                        if (PLC != null && !PLC.userConstant.Any(s => s.name == tempSheet.Cell(row, fieldsUserConstant["Block Name"]).GetValue<string>()))
                        {
                            PLC.userConstant.Add(new userConstant
                            {
                                name = tempSheet.Cell(row, fieldsUserConstant["Block Name"]).GetValue<string>(),
                                type = tempSheet.Cell(row, fieldsUserConstant["Block Type"]).GetValue<string>(),
                                value = tempSheet.Cell(row, fieldsUserConstant["Value"]).GetValue<string>()
                            });
                            message = $"{LogPrefix} Created user constant: {tempSheet.Cell(row, fieldsUserConstant["Block Name"]).GetValue<string>()}";
                            _logger.Information(message);
                        }
                    }
                    message = $"{LogPrefix} Added all support user constants";
                    _logger.Information(message);
                }

                numRowBlockValue = tempSheet.Cell(6, 3).GetValue<string>();
                scopeRowBlockValue = tempSheet.Cell(6, 4).GetValue<string>();
                if (int.TryParse(numRowBlockValue, out numRowBlock) && int.TryParse(scopeRowBlockValue, out scopeRowBlock) && scopeRowBlock > 0)
                {
                    for (int row = numRowBlock + 2; row < numRowBlock + 2 + scopeRowBlock; row++)
                    {
                        var PLC = blocksStruct.FirstOrDefault(plc => plc.namePLC == tempSheet.Cell(row, fieldsExtFCValue["PLCName"]).GetValue<string>());
                        if (PLC != null)
                        {
                            var fc = PLC.dataFC.FirstOrDefault(f => f.name == tempSheet.Cell(row, fieldsExtFCValue["FC"]).GetValue<string>());
                            if (fc != null && !fc.dataExtFCValue.Any(value => value.name == tempSheet.Cell(row, fieldsExtFCValue["Name"]).GetValue<string>()))
                            {
                                fc.dataExtFCValue.Add(new dataExtFCValue
                                {
                                    name = tempSheet.Cell(row, fieldsExtFCValue["Name"]).GetValue<string>(),
                                    type = tempSheet.Cell(row, fieldsExtFCValue["Type"]).GetValue<string>(),
                                    IO = tempSheet.Cell(row, fieldsExtFCValue["IO"]).GetValue<string>(),
                                    Comments = tempSheet.Cell(row, fieldsExtFCValue["Comments"]).GetValue<string>()
                                });
                            }
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
                return true;
            }
            catch (Exception ex)
            {
                message = $"{LogPrefix} Error in excel data processing: {ex.Message}";
                _logger.Error(ex, message);
                return false;
            }
        }
    }
}
