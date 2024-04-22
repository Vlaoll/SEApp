using System;
using System.Diagnostics;
using System.Collections.Generic;
using System.Linq;
using System.Text;
//using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;
using Microsoft.Win32;
using Microsoft.Office.Interop.Excel;
using System.Windows.Shapes;

//using Siemens.Engineering.HW;
//using System.Diagnostics.Eventing.Reader;
//using Siemens.Engineering.SW.Blocks.Interface;
//using System.Xml.Linq;

//using MahApps.Metro.IconPacks;
//using System.Windows.Shapes;
//using Siemens.Engineering.SW.Blocks;

namespace TIAApp
{
    public class ExcelPrj
    {
        private string excelPath = string.Empty;
        private string message = string.Empty;
        private Excel.Application oXL;
        private Excel._Workbook wb;
        private Excel._Worksheet mainWorkSheet;
        private Excel._Worksheet eqWorkSheet;
        private List<Excel._Worksheet> worksheets;



        private List<dataPLC> blocksStruct = new List<dataPLC>();
        public string Message
        {
            get { return message; }

        }
        public List<dataPLC> BlocksStruct
        {
            get { return blocksStruct; }


        }

        Dictionary<string, int> fields = new Dictionary<string, int>
        {
            {"Status", 1},
            {"PLCName", 2},
            {"PrjNum", 3},
            {"EqName", 4},
            {"EqArea", 5 },
            {"EqComments", 6 },
            {"EqType", 7},
            {"PLCNumber", 8},
            {"InstanceOfName", 9},            
            {"GroupDB", 10},
            {"NameFC", 11},
            {"Variant", 12},
            {"ObjName", 13},
            {"PicNum", 14},
            {"ObjTagName", 15},
            {"HMIType", 16},
            {"TypicalPDL", 17},
            {"WorkPDL", 18},
            {"X", 19},
            {"Y", 20},
            {"Width", 21},
            {"Height", 22},
            {"ScaleMsg", 23},
            {"Addition_01", 24},
            {"Addition_02", 25},
            {"Addition_03", 26},
            {"Addition_04", 27},
            {"Addition_05", 28},
            {"Addition_06", 29},
            {"Addition_07", 30},
            {"Addition_08", 31},
            {"Addition_09", 32},
            {"Addition_10", 33},
        };
        Dictionary<string, int> fieldsBlockData = new Dictionary<string, int>
        {
            {"Block Name", 2},
            {"Block Number", 3},
            {"Block Group", 4},
            {"PLCName", 5},
            {"Block Type", 6}
        };
        Dictionary<string, int> fieldsSupportDB = new Dictionary<string, int>
        {
            {"Block Name", 2},            
            {"Block Group", 3},
            {"PLCName", 4},
            {"Block Type", 5},
            {"Path", 6},
            {"Types", 7}
        };
        Dictionary<string, int> fieldsUserConstant = new Dictionary<string, int>
        {
            {"Block Name", 2},
            {"Block Type", 3},
            {"Value", 4},
            {"PLCName", 5}
        };
        Dictionary<string, int> fieldsExtFCValue = new Dictionary<string, int>
        {
            {"Name", 2},
            {"Type", 3},
            {"IO", 4},
            {"Comments", 5},
            {"FC", 6},
            {"PLCName", 7}
        };

        public ExcelPrj()
        {
            oXL = new Excel.Application();
            
            //Excel.Visible = true;
        }


        public string SearchProject(string filter = "Excel |*.xlsx;*.xlsm")
        {
            OpenFileDialog fileSearch = new OpenFileDialog();

            fileSearch.Multiselect = false;
            fileSearch.ValidateNames = true;
            fileSearch.DereferenceLinks = false; // Will return .lnk in shortcuts.
            fileSearch.Filter = filter;            
            fileSearch.RestoreDirectory = true;
            fileSearch.InitialDirectory = Environment.CurrentDirectory;            

            fileSearch.ShowDialog();

            string ProjectPath = fileSearch.FileName.ToString();

            if (string.IsNullOrEmpty(ProjectPath) == false)
            {
                message = "[Excel]" + "Opening excel file " + ProjectPath;
                Trace.WriteLine(message);
                OpenExcelFile(ProjectPath);                
            }
            return ProjectPath;
        }
        public void OpenExcelFile(string filename, string mainSheetName = "Main")
        //Open Excel file and create sheets list and set main sheet 
        {
            try
            {
                excelPath = new FileInfo(filename).FullName;                
                wb = (Excel._Workbook)(oXL.Workbooks.Open(excelPath));
                worksheets = new List<Excel._Worksheet>();

                foreach (Excel._Worksheet item in wb.Worksheets)
                {
                    worksheets.Add(item);
                }

                mainWorkSheet = worksheets.Where(ws => ws.Name == mainSheetName).FirstOrDefault();
                message = "[Excel]" +  "Master Excel " + excelPath + " is opened";
                Trace.WriteLine(message);
            }
            catch (Exception ex)
            { 

                message = "[Excel]" +  "Error while opening excel" + ex.Message;               
                Trace.WriteLine(message);                
            }
        }
        public void CloseExcelFile(bool save = false)
        {
            try
            {
                if (save){wb.Save();}                
                wb.Close(false);
                oXL.Quit();                
                message = "[Excel]" +  "Master Excel " + excelPath + " is closed";
                Trace.WriteLine(message);
            }
            catch (Exception ex)
            {
                message = "[Excel]" +  "Error while closing excel" + ex.Message;
                Trace.WriteLine(message);
            } 

            System.Diagnostics.Process[] process = System.Diagnostics.Process.GetProcessesByName("Excel");
            foreach (System.Diagnostics.Process p in process)
            {
                if (!string.IsNullOrEmpty(p.ProcessName))
                {
                    try { p.Kill(); }
                    catch { }
                }
            }
        }
        public bool ReadExcelData(string status)
        {
            try 
            {
                Excel.Range last = mainWorkSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing);

                List<dataBlock> dataBlock = new List<dataBlock>();
                Excel.Range rng = (Excel.Range)mainWorkSheet.Range[mainWorkSheet.Cells[3, 1], mainWorkSheet.Cells[last.Row, fields.Count]];
                
                object[,] cellValues = (object[,])rng.Value2;

                List<string> lst = cellValues.Cast<object>()
                             .Select(o => o != null?o.ToString():"")
                             .ToList();
                for (int i = 0; i < (last.Row-2)* fields.Count ;)
                {
                    if (lst[i].Contains(status))
                    {
                        string eqName = string.Empty;
                        eqWorkSheet = null;
                        string Status = lst[i];
                        string PLC = lst[i + fields["PLCName"] - 1];
                        string typeEq =  lst[i + fields["EqType"] - 1];
                        string[] tempV;

                        //Add new PLC data to DB
                        if (blocksStruct.Where(item => item.namePLC == PLC).Count() == 0)
                        {
                            blocksStruct.Add(new dataPLC()
                            {
                                namePLC = PLC,
                                Equipment = new List<dataEq>(),
                                instanceDB = new List<dataBlock>(),
                                dataFC = new List<dataFunction>(),
                                dataSupportBD = new List<dataSupportBD>(),
                                userConstant = new List<userConstant>(),
                            });
                            message = "[Excel]" + "Created structure for PLC: " + PLC;
                            Trace.WriteLine(message);
                        }
                        dataPLC tempItem = blocksStruct.Where(item => item.namePLC == PLC).First();

                        //check overload equpment list
                        if (tempItem.instanceDB.Count > 250)
                        {
                            i = i + fields.Count;
                            continue;
                        }    

                        //Searching type's sheet for equipment                       
                        eqWorkSheet = worksheets.Where(ws => ws.Name.Equals(typeEq, StringComparison.OrdinalIgnoreCase)).FirstOrDefault();
                        if (eqWorkSheet == null)
                        {
                            message = "[Excel]" + "Doesn't exist sheet with settings for: " + typeEq;
                            Trace.WriteLine(message);
                            i = i + fields.Count;
                            break;
                        }

                        eqName = eqWorkSheet.Name;

                        //Create types equipment in list, if  it missing                        
                        if (tempItem.Equipment.Where(item => item.typeEq == typeEq).Count() == 0)
                        {
                            
                            tempItem.Equipment.Add(new dataEq()
                            {
                                typeEq = typeEq,
                                isExtended = eqWorkSheet.Cells[2, 5].Value2 == "Extended" ? true : false,
                                FB = new List<dataLibrary>(),
                                dataTag = new List<dataTag>(),
                                dataParameter = new List<dataParameter>(),
                                dataExtSupportBlock = new List<dataExtSupportBlock>(),
                                dataConstant = new List<userConstant>(),
                                dataSupportBD = new List<dataSupportBD>(),
                                dataDataBlockValue = new List<dataDataBlockValue>(),
                            });
                            message = "[Excel]" + "Created structure for equipment type: " + typeEq;
                            Trace.WriteLine(message);

                            dataEq tempEq = tempItem.Equipment.Where(item => item.typeEq == typeEq).First();

                            //Add all support FB and UDT - need to review
                            int numRowBlock = Convert.ToInt32(eqWorkSheet.Cells[10, 3].Value2);
                            int scopeRowBlock = Convert.ToInt32(eqWorkSheet.Cells[10, 4].Value2);
                            if (scopeRowBlock != 0)
                            {
                                try
                                {
                                    for (int row = numRowBlock + 2; row < numRowBlock + 2 + scopeRowBlock; row++)
                                    {

                                        if (!tempEq.FB.Any(item => item.name == eqWorkSheet.Cells[row, 2].Value2))
                                        {
                                            tempEq.FB.Add(new dataLibrary()
                                            {
                                                name = eqWorkSheet.Cells[row, 2].Value2,
                                                group = eqWorkSheet.Cells[row, 3].Value2,
                                                path = eqWorkSheet.Cells[row, 4].Value2,
                                                isType = string.IsNullOrEmpty(eqWorkSheet.Cells[row, 5].Value2) ? false : eqWorkSheet.Cells[row, 5].Value2.ToLower() == "yes" ? true : false,

                                            });
                                        }
                                    }
                                    message = "[Excel]" + "Added all support FB and UDT for equipment type: " + typeEq;
                                    Trace.WriteLine(message);
                                }
                                catch (Exception)
                                {

                                    message = "[Excel:Error]" + "Wrong settings  support FB and UDT for equipment type: " + typeEq;
                                    Trace.WriteLine(message);
                                }
                                
                            }

                            //Add all needed IO signals (binary, analog)
                            numRowBlock = Convert.ToInt32(eqWorkSheet.Cells[11, 3].Value2);
                            scopeRowBlock = Convert.ToInt32(eqWorkSheet.Cells[11, 4].Value2);
                            if (scopeRowBlock != 0)
                            {
                                if (tempEq.dataTag.Count == 0)
                                {
                                    try
                                    {
                                        for (int row = numRowBlock + 2; row < numRowBlock + 2 + scopeRowBlock; row++)
                                        {
                                            string tAdress = "";
                                            if (!string.IsNullOrEmpty(eqWorkSheet.Cells[row, 4].Value2))
                                            {
                                                switch (eqWorkSheet.Cells[row, 4].Value2.ToLower())
                                                {
                                                    case "bool":
                                                        tAdress = "M100.0";
                                                        break;
                                                    case "word" :
                                                        tAdress = "MW1000";
                                                        break;
                                                    case "int":
                                                        tAdress = "MW1000";
                                                        break;

                                                    default:
                                                        tAdress = "";
                                                        break;
                                                }
                                            }
                                            
                                            tempEq.dataTag.Add(new dataTag()
                                            {
                                                name = Convert.ToString(eqWorkSheet.Cells[row, 2].Value2),
                                                link = Convert.ToString(eqWorkSheet.Cells[row, 3].Value2),
                                                type = Convert.ToString(eqWorkSheet.Cells[row, 4].Value2),
                                                adress = tAdress,
                                                table = string.IsNullOrEmpty(eqWorkSheet.Cells[row, 5].Value2) ? "@Eq_IOTable" : Convert.ToString(eqWorkSheet.Cells[row, 5].Value2),
                                                comment = Convert.ToString(eqWorkSheet.Cells[row, 6].Value2),
                                                variant = new List<string>(),
                                            }) ;
                                            tempV = Array.Empty<string>();
                                            tempV = eqWorkSheet.Cells[row, 7].Value2 is null ? Array.Empty<string>() : Convert.ToString(eqWorkSheet.Cells[row, 7].Value2).Split(',');
                                            if (tempV.Length > 0)
                                            {
                                                foreach (string v in tempV)
                                                {
                                                    tempEq.dataTag.Last().variant.Add(v);
                                                }
                                            }
                                        }
                                        message = "[Excel]" + "Added all needed IO signals (binary, analog) for equipment type: " + typeEq;
                                        Trace.WriteLine(message);
                                    }
                                    catch (Exception)
                                    {

                                        message = "[Excel:Error]" + "Wrong settings IO signals (binary, analog) for equipment type: " + typeEq;
                                        Trace.WriteLine(message);
                                    }
                                    
                                }
                            }

                            //Add all needed parameters
                            numRowBlock = Convert.ToInt32(eqWorkSheet.Cells[12, 3].Value2);
                            scopeRowBlock = Convert.ToInt32(eqWorkSheet.Cells[12, 4].Value2);
                            if (scopeRowBlock != 0)
                            {
                                if (tempEq.dataParameter.Count == 0)
                                {
                                    try
                                    {
                                        for (int row = numRowBlock + 2; row < numRowBlock + 2 + scopeRowBlock; row++)
                                        {
                                            tempEq.dataParameter.Add(new dataParameter()
                                            {
                                                name = eqWorkSheet.Cells[row, 2].Value2,
                                                type = eqWorkSheet.Cells[row, 3].Value2,
                                                link = eqWorkSheet.Cells[row, 4].Value2,
                                            });
                                        }
                                        message = "[Excel]" + "Added all needed parameters for equipment type: " + typeEq;
                                        Trace.WriteLine(message);
                                    }
                                    catch (Exception)
                                    {

                                        message = "[Excel:Error]" + "Wrong settings parameters for equipment type: " + typeEq;
                                        Trace.WriteLine(message);
                                    }
                                    
                                }
                            }

                            //Add all needed support DB
                            numRowBlock = Convert.ToInt32(eqWorkSheet.Cells[13, 3].Value2);
                            scopeRowBlock = Convert.ToInt32(eqWorkSheet.Cells[13, 4].Value2);
                            if (scopeRowBlock != 0)
                            {
                                if (tempEq.dataExtSupportBlock.Count == 0)
                                {
                                    try
                                    {
                                        for (int row = numRowBlock + 2; row < numRowBlock + 2 + scopeRowBlock; row++)
                                        {
                                            tempEq.dataExtSupportBlock.Add(new dataExtSupportBlock()
                                            {

                                                name = eqWorkSheet.Cells[row, 2].Value2,
                                                instanceOfName = eqWorkSheet.Cells[row, 3].Value2,
                                                number = Convert.ToInt32(eqWorkSheet.Cells[row, 4].Value2),
                                                variant = new List<string>(),
                                            });
                                            tempV = Array.Empty<string>();
                                            tempV = eqWorkSheet.Cells[row, 5].Value2 is null ? Array.Empty<string>(): Convert.ToString(eqWorkSheet.Cells[row, 5].Value2).Split(',');
                                            if (tempV.Length>0)
                                            {
                                                foreach (string v in tempV)
                                                {
                                                    tempEq.dataExtSupportBlock.Last().variant.Add(v);
                                                }
                                            }
                                            
                                        }
                                        message = "[Excel]" + "Added all needed support DB for equipment type: " + typeEq;
                                        Trace.WriteLine(message);
                                    }
                                    catch (Exception)
                                    {

                                        message = "[Excel]" + "Added all needed support DB for equipment type: " + typeEq;
                                        Trace.WriteLine(message);
                                    }
                                    
                                }
                            }

                            //Add all needed Constants
                            numRowBlock = Convert.ToInt32(eqWorkSheet.Cells[14, 3].Value2);
                            scopeRowBlock = Convert.ToInt32(eqWorkSheet.Cells[14, 4].Value2);
                            if (scopeRowBlock != 0)
                            {
                                if (tempEq.dataConstant.Count == 0)
                                {
                                    try
                                    {
                                        for (int row = numRowBlock + 2; row < numRowBlock + 2 + scopeRowBlock; row++)
                                        {
                                            tempEq.dataConstant.Add(new userConstant()
                                            {

                                                name = eqWorkSheet.Cells[row, 2].Value2,
                                                type = eqWorkSheet.Cells[row, 3].Value2,
                                                value = eqWorkSheet.Cells[row, 4].Value2,
                                                table = eqWorkSheet.Cells[row, 5].Value2,
                                            });
                                        }
                                        message = "[Excel]" + "Added all needed constants equipment type: " + typeEq;
                                        Trace.WriteLine(message);
                                    }
                                    catch (Exception)
                                    {

                                        message = "[Excel:Error]" + "Wrong settings constants equipment type: " + typeEq;
                                        Trace.WriteLine(message);
                                    }
                                    
                                    
                                }
                            }

                            //Add all needed DBs
                            numRowBlock = Convert.ToInt32(eqWorkSheet.Cells[15, 3].Value2);
                            scopeRowBlock = Convert.ToInt32(eqWorkSheet.Cells[15, 4].Value2);
                            if (scopeRowBlock != 0)
                            {
                                if (tempEq.dataSupportBD.Count == 0)
                                {
                                    try
                                    {
                                        for (int row = numRowBlock + 2; row < numRowBlock + 2 + scopeRowBlock; row++)
                                        {                                            
                                            tempEq.dataSupportBD.Add(new dataSupportBD()
                                            {

                                                name = eqWorkSheet.Cells[row, 2].Value2,
                                                number = Convert.ToInt32(eqWorkSheet.Cells[row, 3].Value2),
                                                group = eqWorkSheet.Cells[row, 4].Value2,
                                                path = eqWorkSheet.Cells[row, 5].Value2,
                                                isType = string.IsNullOrEmpty(eqWorkSheet.Cells[row, 6].Value2) ? false : eqWorkSheet.Cells[row, 6].Value2.ToLower() == "yes" ? true : false,
                                                isRetain = string.IsNullOrEmpty(eqWorkSheet.Cells[row, 7].Value2) ? false : eqWorkSheet.Cells[row, 7].Value2.ToLower() == "yes" ? true : false,
                                                isOptimazed = string.IsNullOrEmpty(eqWorkSheet.Cells[row, 8].Value2) ? false : eqWorkSheet.Cells[row, 8].Value2.ToLower() == "yes" ? true : false,
                                            });
                                        }
                                        message = "[Excel]" + "Added all needed support DB for equipment type: " + typeEq;
                                        Trace.WriteLine(message);
                                    }
                                    catch (Exception)
                                    {

                                        message = "[Excel:Error]" + "Wrong settings support DB for equipment type: " + typeEq;
                                        Trace.WriteLine(message);
                                    }
                                    
                                }
                            }

                            //Add all needed DB's values
                            numRowBlock = Convert.ToInt32(eqWorkSheet.Cells[16, 3].Value2);
                            scopeRowBlock = Convert.ToInt32(eqWorkSheet.Cells[16, 4].Value2);
                            if (scopeRowBlock != 0)
                            {
                                if (tempEq.dataDataBlockValue.Count == 0)
                                {
                                    try
                                    {
                                        for (int row = numRowBlock + 2; row < numRowBlock + 2 + scopeRowBlock; row++)
                                        {
                                            tempEq.dataDataBlockValue.Add(new dataDataBlockValue()
                                            {

                                                name = eqWorkSheet.Cells[row, 2].Value2,
                                                type = eqWorkSheet.Cells[row, 3].Value2,
                                                DB = eqWorkSheet.Cells[row, 4].Value2,
                                            });
                                        }
                                        message = "[Excel]" + "Added all needed DB's values for equipment type: " + typeEq;
                                        Trace.WriteLine(message);
                                    }
                                    catch (Exception)
                                    {

                                        message = "[Excel:Error]" + "Wrong settings DB's values for equipment type: " + typeEq;
                                        Trace.WriteLine(message);
                                    }
                                    
                                }
                            }
                        }

                        //Check current status of equipment
                        bool isExtended = tempItem.Equipment.Where(item => item.typeEq == typeEq).First().isExtended;
                        tempItem.instanceDB.Add(new dataBlock(){
                            name = isExtended ? "FC_" + lst[i + fields["EqName"] - 1] : "iDB-" + lst[i + fields["EqType"] - 1]  + "|" + lst[i + fields["EqName"] - 1],
                            comment = lst[i + fields["EqComments"] - 1],
                            area = lst[i + fields["EqArea"] - 1],
                            instanceOfName = lst[i + fields["InstanceOfName"] - 1],
                            number = Convert.ToInt32(lst[i + fields["PLCNumber"] - 1]),
                            group = lst[i + fields["GroupDB"] - 1] ,
                            nameFC = lst[i + fields["NameFC"] - 1] ,
                            typeEq = lst[i + fields["EqType"] - 1] ,
                            nameEq = lst[i + fields["EqName"] - 1] ,
                            variant = new List<string>(),
                            excelData = new List<excelData>()
                        });
                        tempV = Array.Empty<string>();
                        tempV = lst[i + fields["Variant"] - 1] == "" ? Array.Empty<string>() : lst[i + fields["Variant"] - 1].Replace('.',',').Split(',');
                        if (tempV.Length > 0)
                        {
                            foreach (string v in tempV)
                            {
                                tempItem.instanceDB.Last().variant.Add(v);
                            }
                        }
                        //Fill all data for equipment
                        var ex = tempItem.instanceDB
                                .Where(inst => inst.nameEq == lst[i + fields["EqName"] - 1])
                                .First()
                                .excelData;
                        foreach (var data in fields)
                        {
                            ex.Add(new excelData()
                            {
                                name = data.Key,
                                column = data.Value,
                                value = lst[i++],
                            });
                        }
                        message = "[Excel]" + "Created Instance for:" + ex.Where(item=>item.name == "EqName").FirstOrDefault().value;
                        Trace.WriteLine(message);
                    }
                    else
                    {
                        i = i + fields.Count;
                    } 
                }
                return true;
            }
            catch (Exception ex)
            {
                message = "[Excel]" + "Error Data in excel  " + ex.Message;
                Trace.WriteLine(message);
                return false;
            }
        }              
        public bool CreateFCList(string sheetBlockDataNane = "PLCData")
        {
            try
            {
                Excel._Worksheet tempSheet = worksheets.Where(ws => ws.Name == sheetBlockDataNane).FirstOrDefault();

                //Add all functions
                int numRowBlock = Convert.ToInt32(tempSheet.Cells[3, 3].Value2);
                int scopeRowBlock = Convert.ToInt32(tempSheet.Cells[3, 4].Value2);
                if (scopeRowBlock != 0)
                {
                    for (int row = numRowBlock + 2; row < numRowBlock + 2 + scopeRowBlock; row++)
                    {
                        var PLC = blocksStruct.Where(plc => plc.namePLC == tempSheet.Cells[row, fieldsBlockData["PLCName"]].Value2);
                        if (PLC.Count() != 0)
                        {
                            if (!PLC.First().dataFC.Any(fc => fc.name == tempSheet.Cells[row, fieldsBlockData["Block Name"]].Value2))
                            {
                                PLC.First().dataFC.Add(new dataFunction()
                                {
                                    name = tempSheet.Cells[row , fieldsBlockData["Block Name"]].Value2,
                                    number = Convert.ToInt32(tempSheet.Cells[row, fieldsBlockData["Block Number"]].Value2),
                                    group = tempSheet.Cells[row, fieldsBlockData["Block Group"]].Value2,
                                    code = new StringBuilder("FUNCTION " + "\"" + tempSheet.Cells[row, fieldsBlockData["Block Name"]].Value2 + "\" : Void\r\n")
                                             .AppendLine("{ S7_Optimized_Access := 'TRUE' }")
                                             .AppendLine("AUTHOR : SE")
                                             .AppendLine("FAMILY : Constructor"),
                                             //.AppendLine("BEGIN"),
                                    dataExtFCValue = new List<dataExtFCValue>()
                                });
                                message = "[Excel]" +  "Created template for function block: " + tempSheet.Cells[row, fieldsBlockData["Block Name"]].Value2;
                                Trace.WriteLine(message);
                            }
                        }
                    }
                    message = "[Excel]" +  "Added all support FCs " ;
                    Trace.WriteLine(message);
                }

                //Add all support BD
                numRowBlock = Convert.ToInt32(tempSheet.Cells[4, 3].Value2);
                scopeRowBlock = Convert.ToInt32(tempSheet.Cells[4, 4].Value2);
                if (scopeRowBlock != 0)
                {
                    for (int row = numRowBlock + 2; row < numRowBlock + 2 + scopeRowBlock; row++)
                    {
                        var PLC = blocksStruct.Where(plc => plc.namePLC == tempSheet.Cells[row, fieldsSupportDB["PLCName"]].Value2);
                        if (PLC.Count() != 0)
                        {
                            if (!PLC.First().dataSupportBD.Any(s => s.name == tempSheet.Cells[row, fieldsSupportDB["Block Name"]].Value2))
                            {
                                PLC.First().dataSupportBD.Add(new dataSupportBD()
                                {
                                    name = tempSheet.Cells[row, fieldsSupportDB["Block Name"]].Value2,                                    
                                    group = tempSheet.Cells[row, fieldsSupportDB["Block Group"]].Value2,
                                    type = tempSheet.Cells[row, fieldsSupportDB["Block Type"]].Value2,
                                    path = tempSheet.Cells[row, fieldsSupportDB["Path"]].Value2 ,
                                    isType = string.IsNullOrEmpty(tempSheet.Cells[row, fieldsSupportDB["Types"]].Value2) ? false : tempSheet.Cells[row, fieldsSupportDB["Types"]].Value2.ToLower() == "yes" ? true : false,
                                    isRetain = false,
                                    
                                });
                                message = "[Excel]" +  "Created support BD in DB: " + tempSheet.Cells[row, fieldsSupportDB["Block Name"]].Value2;
                                Trace.WriteLine(message);
                            }
                        }
                    }
                    message = "[Excel]" +  "Added all support BD. ";
                    Trace.WriteLine(message);
                }

                //Add all support user constants
                numRowBlock = Convert.ToInt32(tempSheet.Cells[5, 3].Value2);
                scopeRowBlock = Convert.ToInt32(tempSheet.Cells[5, 4].Value2);
                if (scopeRowBlock != 0)
                {
                    for (int row = numRowBlock + 2; row < numRowBlock + 2 + scopeRowBlock; row++)
                    {
                        var PLC = blocksStruct.Where(plc => plc.namePLC == tempSheet.Cells[row, fieldsUserConstant["PLCName"]].Value2);
                        if (PLC.Count() != 0)
                        {
                            if (!PLC.First().userConstant.Any(s => s.name == tempSheet.Cells[row, fieldsUserConstant["Block Name"]].Value2))
                            {
                                PLC.First().userConstant.Add(new userConstant()
                                {
                                    name = tempSheet.Cells[row, fieldsUserConstant["Block Name"]].Value2,                                   
                                    type = tempSheet.Cells[row, fieldsUserConstant["Block Type"]].Value2,
                                    value = Convert.ToString(tempSheet.Cells[row, fieldsUserConstant["Value"]].Value2) ,

                                });
                                message = "[Excel]" +  "Created user constant: " + tempSheet.Cells[row, fieldsUserConstant["Block Name"]].Value2;
                                Trace.WriteLine(message);
                            }
                        }
                    }
                    message = "[Excel]" +  "Added all support user constants. ";
                    Trace.WriteLine(message);
                }

                //Add all extended values for functions
                numRowBlock = Convert.ToInt32(tempSheet.Cells[6, 3].Value2);
                scopeRowBlock = Convert.ToInt32(tempSheet.Cells[6, 4].Value2);
                if (scopeRowBlock != 0)
                {
                    for (int row = numRowBlock + 2; row < numRowBlock + 2 + scopeRowBlock; row++)
                    {
                        
                        var PLC = blocksStruct.Where(plc => plc.namePLC == tempSheet.Cells[row, fieldsExtFCValue["PLCName"]].Value2);
                        if (PLC.Count() != 0)
                        {
                            if (PLC.First().dataFC.Any(fc => fc.name == tempSheet.Cells[row, fieldsExtFCValue["FC"]].Value2))
                            {
                                var tempFC = PLC.First().dataFC.Where(fc => fc.name == tempSheet.Cells[row, fieldsExtFCValue["FC"]].Value2).First().dataExtFCValue;
                                if (!tempFC.Any(value => value.name == tempSheet.Cells[row, fieldsExtFCValue["Name"]].Value2))
                                {
                                    tempFC.Add(new dataExtFCValue() 
                                    { 
                                        name = tempSheet.Cells[row, fieldsExtFCValue["Name"]].Value2,
                                        type = tempSheet.Cells[row, fieldsExtFCValue["Type"]].Value2,
                                        IO = tempSheet.Cells[row, fieldsExtFCValue["IO"]].Value2,
                                        Comments = tempSheet.Cells[row, fieldsExtFCValue["Comments"]].Value2,
                                    });
                                }                                
                            }
                        }
                    }
                    foreach (var PLC in BlocksStruct)
                    {
                        foreach (var FC in PLC.dataFC)
                        {
                            List<dataExtFCValue> list = FC.dataExtFCValue.Where(type => type.IO == "I").ToList();
                            if (list.Count > 0)
                            {
                                FC.code.AppendLine("VAR_INPUT");
                                foreach (var values in list)
                                {
                                    FC.code.AppendLine("\"" + values.name + "\" : " + values.type + ";   // " + values.Comments);
                                }
                                FC.code.AppendLine("END_VAR");
                            }
                            list = FC.dataExtFCValue.Where(type => type.IO == "O").ToList();
                            if (list.Count > 0)
                            {
                                FC.code.AppendLine("VAR_OUTPUT");
                                foreach (var values in list)
                                {
                                    FC.code.AppendLine("\"" + values.name + "\" : " + values.type + ";   // " + values.Comments);
                                }
                                FC.code.AppendLine("END_VAR");
                            }
                            list = FC.dataExtFCValue.Where(type => type.IO == "IO").ToList();
                            if (list.Count > 0)
                            {
                                FC.code.AppendLine("VAR_IN_OUT");
                                foreach (var values in list)
                                {
                                    FC.code.AppendLine("\"" + values.name + "\" : " + values.type + ";   // " + values.Comments);
                                }
                                FC.code.AppendLine("END_VAR");
                            }
                            list = FC.dataExtFCValue.Where(type => type.IO == "T").ToList();
                            if (list.Count > 0)
                            {
                                FC.code.AppendLine("VAR_TEMP");
                                foreach (var values in list)
                                {
                                    FC.code.AppendLine("\"" + values.name + "\" : " + values.type + ";   // " + values.Comments);
                                }
                                FC.code.AppendLine("END_VAR");
                            }
                            list = FC.dataExtFCValue.Where(type => type.IO == "C").ToList();
                            if (list.Count > 0)
                            {
                                FC.code.AppendLine("VAR CONSTANT");
                                foreach (var values in list)
                                {
                                    FC.code.AppendLine("\"" + values.name + "\" : " + values.type + ";   // " + values.Comments);
                                }
                                FC.code.AppendLine("END_VAR");
                            }
                            FC.code.AppendLine("BEGIN");
                            message = "[Excel]" + "extended values for function: " + FC.name ;
                            Trace.WriteLine(message);
                        }
                        
                    }
                }
                return true;
            }
            catch (Exception ex)
            {
                message = "[Excel]" +  "Error Data in excel  " + ex.Message;
                Trace.WriteLine(message);
                return false;
            }
        }




        //public void CreateTagExcel(string sourceTagPath, bool close = false)
        //{
        //    try
        //    {
        //        excelPath = new FileInfo(sourceTagPath).FullName;
        //        object misValue = System.Reflection.Missing.Value;
        //        wb = oXL.Workbooks.Add(misValue);
        //        var xlSheets = wb.Sheets as Excel.Sheets;
        //        mainWorkSheet = (Excel._Worksheet)xlSheets.Add(xlSheets[1], Type.Missing, Type.Missing, Type.Missing);
        //        mainWorkSheet.Name = "PLC Tags";

        //        mainWorkSheet.Cells[1, 1].Value = "Name";
        //        mainWorkSheet.Cells[1, 2].Value = "Path";
        //        mainWorkSheet.Cells[1, 3].Value = "Data Type";
        //        mainWorkSheet.Cells[1, 4].Value = "Logical Address";
        //        mainWorkSheet.Cells[1, 5].Value = "Comment";
        //        mainWorkSheet.Cells[1, 6].Value = "Hmi Visible";
        //        mainWorkSheet.Cells[1, 7].Value = "Hmi Accessible";
        //        mainWorkSheet.Cells[1, 8].Value = "Hmi Writeable";
        //        mainWorkSheet.Cells[1, 9].Value = "Typeobject ID";
        //        mainWorkSheet.Cells[1, 10].Value = "Version ID";

        //        if (File.Exists(excelPath)) { File.Delete(excelPath); }
        //        wb.SaveAs(excelPath);
        //        if (close) { wb.Close(); }

        //        message = "[Excel]" + "Created new excel file for PLC Tags ";
        //        Trace.WriteLine(message);
        //    }
        //    catch (Exception ex)
        //    {

        //        message = "[Excel]" + "Error while opening excel" + ex.Message;
        //        Trace.WriteLine(message);
        //    }

        //}
        //public void AddTagToExcel(string name, string path, string type, string address, string comment, string hmiVisible = "True", string hmiAccessible = "True", string hmiWriteable = "True", string typeObj = "",string version = "", bool lastRow = false)
        //{
        //    Excel.Range last = mainWorkSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing);
        //    var row = last.Row + 1;
        //    mainWorkSheet.Cells[row, 1].Value = name;
        //    mainWorkSheet.Cells[row, 2].Value = path;
        //    mainWorkSheet.Cells[row, 3].Value = type;
        //    mainWorkSheet.Cells[row, 4].Value = address;
        //    mainWorkSheet.Cells[row, 5].Value = comment;
        //    mainWorkSheet.Cells[row, 6].Value = hmiVisible;
        //    mainWorkSheet.Cells[row, 7].Value = hmiAccessible;
        //    mainWorkSheet.Cells[row, 8].Value = hmiWriteable;
        //    mainWorkSheet.Cells[row, 9].Value = typeObj;
        //    mainWorkSheet.Cells[row, 10].Value = version;
        //    if (lastRow)
        //    {
        //        wb.Save();
        //        wb.Close();
        //    }
        //}

        //public System.Data.DataTable WorkbookToDataTable()
        //{
        //    Excel.Range last = mainWorkSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing);

        //    List<string> newRow = new List<string>();
        //    var dataTable = new System.Data.DataTable();

        //    for (int rowNum = 3; rowNum <= last.Row; rowNum++)
        //    {
        //        for (int columnNum = 1; columnNum <= 32; columnNum++)
        //        {
        //            // In my solution, the first row of the table is assumed to be header rows.
        //            // So the first row's items will be the name of each column
        //            if (rowNum == 1)
        //            {
        //                dataTable.Columns.Add(new System.Data.DataColumn(mainWorkSheet.Cells[rowNum, columnNum].Value2.ToString(), typeof(object)));
        //            }
        //            else if (mainWorkSheet.Cells[rowNum, columnNum].Value2 == "Blocks")
        //            {
        //                newRow.Add(mainWorkSheet.Cells[rowNum, columnNum].Value2.ToString());
        //            }
        //        }
        //        if (rowNum != 1)
        //        {
        //            dataTable.Rows.Add(newRow);
        //            newRow = new List<string>();
        //        }
        //    }


        //    return dataTable;
        //}

        //public void CreateInstanceList(string status)
        //{
        //    try
        //    {
        //        Excel.Range last = mainWorkSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing);

        //        //Checking all exist rows in main sheet
        //        for (int i = 3; i < last.Row + 1; i++)
        //        {
        //            string eqName = string.Empty;
        //            eqWorkSheet = null;
        //            string Status = mainWorkSheet.Cells[i, fields["Status"]].Value2;
        //            string PLC = mainWorkSheet.Cells[i, fields["PLCName"]].Value2;
        //            string typeEq = mainWorkSheet.Cells[i, fields["EqType"]].Value2;
        //            //Check current status of equipment
        //            if (Status.Contains(status))
        //            {
        //                //Add new PLC data to DB
        //                if (blocksStruct.Where(item => item.namePLC == PLC).Count() == 0)
        //                {
        //                    blocksStruct.Add(new dataPLC()
        //                    {
        //                        namePLC = PLC,
        //                        Equipment = new List<dataEq>(),
        //                        instanceDB = new List<dataBlock>(),
        //                        dataFC = new List<dataFunction>(),
        //                        dataSupportBD = new List<dataSupportBD>(),
        //                        userConstant = new List<userConstant>(),
        //                    });
        //                    message = "[Excel]" + "Created structure for PLC: " + PLC;
        //                    Trace.WriteLine(message);
        //                }
        //                dataPLC tempItem = blocksStruct.Where(item => item.namePLC == PLC).First();

        //                //Searching type's sheet for equipment
        //                eqWorkSheet = worksheets.Where(ws => ws.Name == typeEq).FirstOrDefault();
        //                if (eqWorkSheet == null)
        //                {
        //                    message = "[Excel]" + "Doesn't exist sheet with settings for: " + typeEq;
        //                    Trace.WriteLine(message);
        //                    break;
        //                }
        //                eqName = eqWorkSheet.Name;

        //                //Create types equipment in list, if  it missing                        
        //                if (tempItem.Equipment.Where(item => item.typeEq == typeEq).Count() == 0)
        //                {
        //                    tempItem.Equipment.Add(new dataEq()
        //                    {
        //                        typeEq = typeEq,
        //                        isExtended = eqWorkSheet.Cells[2, 5].Value2 == "Extended" ? true : false,
        //                        FB = new List<dataLibrary>(),
        //                        dataTag = new List<dataTag>(),
        //                        dataParameter = new List<dataParameter>(),
        //                        dataExtSupportBlock = new List<dataExtSupportBlock>(),
        //                    });
        //                    message = "[Excel]" + "Created structure for equipment type: " + typeEq;
        //                    Trace.WriteLine(message);

        //                    dataEq tempEq = tempItem.Equipment.Where(item => item.typeEq == typeEq).First();

        //                    //Add all support FB and UDT - need to review
        //                    int numRowBlock = Convert.ToInt32(eqWorkSheet.Cells[10, 3].Value2);
        //                    int scopeRowBlock = Convert.ToInt32(eqWorkSheet.Cells[10, 4].Value2);
        //                    if (scopeRowBlock != 0)
        //                    {
        //                        for (int row = numRowBlock + 2; row < numRowBlock + 2 + scopeRowBlock; row++)
        //                        {

        //                            if (!tempEq.FB.Any(item => item.name == eqWorkSheet.Cells[row, 2].Value2))
        //                            {
        //                                tempEq.FB.Add(new dataLibrary()
        //                                {
        //                                    name = eqWorkSheet.Cells[row, 2].Value2,
        //                                    group = eqWorkSheet.Cells[row, 3].Value2,
        //                                    path = eqWorkSheet.Cells[row, 4].Value2
        //                                });
        //                            }
        //                        }
        //                        message = "[Excel]" + "Added all support FB and UDT for equipment type: " + typeEq;
        //                        Trace.WriteLine(message);
        //                    }

        //                    //Add all needed IO signals (binary, analog)
        //                    numRowBlock = Convert.ToInt32(eqWorkSheet.Cells[11, 3].Value2);
        //                    scopeRowBlock = Convert.ToInt32(eqWorkSheet.Cells[11, 4].Value2);
        //                    if (scopeRowBlock != 0)
        //                    {
        //                        if (tempEq.dataTag.Count == 0)
        //                        {
        //                            for (int k = numRowBlock + 2; k < numRowBlock + 2 + scopeRowBlock; k++)
        //                            {
        //                                tempEq.dataTag.Add(new dataTag()
        //                                {
        //                                    name = eqWorkSheet.Cells[k, 2].Value2,
        //                                    link = eqWorkSheet.Cells[k, 3].Value2,
        //                                    type = eqWorkSheet.Cells[k, 4].Value2,
        //                                    adress = (eqWorkSheet.Cells[k, 4].Value2 == "Bool") ? "M100.0" : (eqWorkSheet.Cells[k, 4].Value2 == "Word") ? "MW100" : "",
        //                                    table = (eqWorkSheet.Cells[k, 5].Value2 == null) ? "@Eq_IOTable" : eqWorkSheet.Cells[k, 5].Value2,
        //                                    comment = Convert.ToString(eqWorkSheet.Cells[k, 6].Value2),
        //                                });
        //                            }
        //                            message = "[Excel]" + "Added all needed IO signals (binary, analog) for equipment type: " + typeEq;
        //                            Trace.WriteLine(message);
        //                        }
        //                    }

        //                    //Add all needed parameters
        //                    numRowBlock = Convert.ToInt32(eqWorkSheet.Cells[12, 3].Value2);
        //                    scopeRowBlock = Convert.ToInt32(eqWorkSheet.Cells[12, 4].Value2);
        //                    if (scopeRowBlock != 0)
        //                    {
        //                        if (tempEq.dataParameter.Count == 0)
        //                        {
        //                            for (int k = numRowBlock + 2; k < numRowBlock + 2 + scopeRowBlock; k++)
        //                            {
        //                                tempEq.dataParameter.Add(new dataParameter()
        //                                {
        //                                    name = eqWorkSheet.Cells[k, 2].Value2,
        //                                    type = eqWorkSheet.Cells[k, 3].Value2,
        //                                    link = eqWorkSheet.Cells[k, 4].Value2,
        //                                });
        //                            }
        //                            message = "[Excel]" + "Added all needed parameters for equipment type: " + typeEq;
        //                            Trace.WriteLine(message);
        //                        }
        //                    }

        //                    //Add all needed support DB
        //                    numRowBlock = Convert.ToInt32(eqWorkSheet.Cells[13, 3].Value2);
        //                    scopeRowBlock = Convert.ToInt32(eqWorkSheet.Cells[13, 4].Value2);
        //                    if (scopeRowBlock != 0)
        //                    {
        //                        if (tempEq.dataExtSupportBlock.Count == 0)
        //                        {
        //                            for (int k = numRowBlock + 2; k < numRowBlock + 2 + scopeRowBlock; k++)
        //                            {
        //                                tempEq.dataExtSupportBlock.Add(new dataExtSupportBlock()
        //                                {

        //                                    name = eqWorkSheet.Cells[k, 2].Value2,
        //                                    instanceOfName = eqWorkSheet.Cells[k, 3].Value2,
        //                                    number = Convert.ToInt32(eqWorkSheet.Cells[k, 4].Value2),
        //                                });
        //                            }
        //                            message = "[Excel]" + "Added all needed support DB for equipment type: " + typeEq;
        //                            Trace.WriteLine(message);
        //                        }
        //                    }
        //                }
        //                InstanceDBList(tempItem, i, tempItem.Equipment.Where(item => item.typeEq == typeEq).First().isExtended);

        //            }
        //        }
        //        message = "[Excel]" + "Instance DB List is Ready";
        //        Trace.WriteLine(message);
        //    }
        //    catch (Exception ex)
        //    {
        //        message = "[Excel]" + "Error Data in excel  " + ex.Message;
        //        Trace.WriteLine(message);
        //    }
        //}
        //public bool InstanceDBList(dataPLC itemPLC, int row, bool isExtended)
        //{
        //    try
        //    {
        //        //Add instance element
        //        itemPLC.instanceDB.Add(new dataBlock()
        //        {
        //            name = isExtended ? "FC_" + mainWorkSheet.Cells[row, fields["EqName"]].Value2 : "iDB-" + mainWorkSheet.Cells[row, fields["EqType"]].Value2 + "|" + mainWorkSheet.Cells[row, fields["EqName"]].Value2,
        //            comment = mainWorkSheet.Cells[row, fields["EqComments"]].Value2,
        //            area = mainWorkSheet.Cells[row, fields["EqArea"]].Value2,
        //            instanceOfName = mainWorkSheet.Cells[row, fields["InstanceOfName"]].Value2,
        //            number = Convert.ToInt32(mainWorkSheet.Cells[row, fields["PLCNumber"]].Value2),
        //            group = mainWorkSheet.Cells[row, fields["GroupDB"]].Value2,
        //            nameFC = mainWorkSheet.Cells[row, fields["NameFC"]].Value2,
        //            typeEq = mainWorkSheet.Cells[row, fields["EqType"]].Value2,
        //            nameEq = mainWorkSheet.Cells[row, fields["EqName"]].Value2,
        //            variant = new List<string>(),
        //            excelData = new List<excelData>()
        //        });

        //        //Fill all data for equipment
        //        var ex = itemPLC.instanceDB
        //                .Where(inst => inst.nameEq == mainWorkSheet.Cells[row, fields["EqName"]].Value2)
        //                .First()
        //                .excelData;
        //        foreach (var data in fields)
        //        {
        //            ex.Add(new excelData()
        //            {
        //                name = data.Key,
        //                column = data.Value,
        //                value = Convert.ToString(mainWorkSheet.Cells[row, fields[data.Key]].Value2),
        //            });
        //        }

        //        message = "[Excel]" + "Created Instance for:" + mainWorkSheet.Cells[row, fields["EqName"]].Value2;
        //        Trace.WriteLine(message);
        //        return true;
        //    }
        //    catch (Exception ex)
        //    {
        //        message = "[Excel]" + "Error Data in excel  " + ex.Message;
        //        Trace.WriteLine(message);
        //        return false;
        //    }

        //}

    }
}