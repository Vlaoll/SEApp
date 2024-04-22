using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Diagnostics;
using System.Windows;


using System.IO;

using Siemens.Engineering.SW.Blocks;
using Microsoft.Office.Interop.Excel;
using Siemens.Engineering.Library;
using Siemens.Engineering.SW.Tags;
using System.Data.OleDb;
//using TIAApp.DataModel;
using Excel = Microsoft.Office.Interop.Excel;
using System.Data;

using System.Security.Principal;
using Microsoft.Win32;
using System.Runtime.Remoting.Messaging;
using TIAApp.ViewModels;
using TIAApp.Views;

namespace TIAApp
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    
    public partial class MainWindow : System.Windows.Window
    {
        public MainWindow()
        {
            InitializeComponent();
            
        }

       
    }

}
//private bool CreateSupportBlocks(string ProjectPath, string ProjectLibPath, string libPLC = "libPLC", bool closeProject = false)
//{
//    string filename;
//    try
//    {

//        filename = new FileInfo(ProjectLibPath).FullName;
//        _tiaprj.OpenProject(filename);
//        _tiaprj.ConnectTIA();

//        if (Directory.Exists(exportPath)) Directory.Delete(exportPath, true);
//        if (Directory.Exists(sourcePath)) Directory.Delete(sourcePath, true);

//        Directory.CreateDirectory(exportPath);
//        Directory.CreateDirectory(sourcePath);


//        foreach (dataPLC item in _excelprj.BlocksStruct)
//        {
//            foreach (dataEq equipment in item.Equipment)
//            {
//                foreach (string block in equipment.blockPLC)
//                {
//                    _tiaprj.ExportBlock(libPLC, block, exportPath);
//                }
//            }
//            foreach (dataEq equipment in item.Equipment)
//            {
//                foreach (string block in equipment.udtPLC)
//                {
//                    _tiaprj.GenerateSourceUDT(libPLC, block, sourcePath);
//                }
//            }
//        }
//        _tiaprj.CloseProject();

//        filename = new FileInfo(ProjectPath).FullName;
//        _tiaprj.OpenProject(filename);
//        _tiaprj.ConnectTIA();
//        foreach (dataPLC item in _excelprj.BlocksStruct)
//        {

//            foreach (dataEq equipment in item.Equipment)
//            {
//                foreach (string block in equipment.udtPLC)
//                {
//                    _tiaprj.ImportSource(item.namePLC, block, sourcePath, ".udt");
//                }
//            }
//            _tiaprj.GenerateBlock(item.namePLC);
//            _tiaprj.ClearSource(item.namePLC);
//            _tiaprj.SaveProject();

//            foreach (dataEq equipment in item.Equipment)
//            {
//                foreach (string block in equipment.blockPLC)
//                {
//                    _tiaprj.ImportBlock(item.namePLC, block, exportPath, "@Eq_FB");
//                }
//            }
//            _tiaprj.Compile(item.namePLC);
//        }
//        _tiaprj.SaveProject();

//        if (closeProject) { _tiaprj.CloseProject(); }

//        return true;
//    }
//    catch (Exception ex)
//    {

//         msg = "[General]" + ex.Message;
//        return false;
//    }
//}

//private bool CreateFCCallAllBlocs(bool closeProject = false)
//{
//    StringBuilder io;
//    try
//    {
//        foreach (dataPLC item in _excelprj.BlocksStruct)
//        {
//            _tiaprj.DataPrjListFB.Clear();
//            _tiaprj.CreateListFB(item.namePLC, item.dataListFB);
//            List<dataBlock> orderDataPrjListFB = _tiaprj.DataPrjListFB.OrderBy(order => order.name).ToList();
//            foreach (dataBlock data in orderDataPrjListFB)
//            {
//                io = null;
//                io = new StringBuilder();
//                io.AppendLine("(");
//                if (data.typeEq.Length > 0 && data.nameEq.Length > 0)
//                {
//                    dataEq EqType = _excelprj.BlocksStruct
//                        .Where(plc => plc.namePLC == item.namePLC)
//                        .First()
//                        .Equipment
//                        .Where(eq => eq.typeEq == data.typeEq)
//                        .First();
//                    int numEq = EqType.dataTag.Count;
//                    foreach (dataTag tag in EqType.dataTag)
//                    {
//                        if (numEq == 1)
//                        {
//                            io.Append(tag.io + ":=" + data.nameEq + tag.name);
//                        }
//                        else
//                        {
//                            io.AppendLine(tag.io + ":=" + data.nameEq + tag.name + ",");
//                        }
//                        numEq--;
//                    }
//                }
//                io.Append(");");
//                if (!_excelprj.BlocksStruct.Where(plc => plc.namePLC == item.namePLC).First().dataFC.Any(block => block.name == data.nameFC))
//                {
//                    _excelprj.BlocksStruct.Where(plc => plc.namePLC == item.namePLC).First().dataFC.Add(new dataFunction()
//                    {
//                        name = data.nameFC,
//                        number = data.numFC,
//                        group = data.groupFC,
//                        code = new StringBuilder("FUNCTION " + "\"" + data.nameFC + "\" : Void\r\n")
//                                    .AppendLine("{ S7_Optimized_Access := 'TRUE' }")
//                                    .AppendLine("AUTHOR : SE")
//                                    .AppendLine("FAMILY : Constructor")
//                                    .AppendLine("BEGIN")
//                    });
//                }
//                    _excelprj.BlocksStruct.Where(plc => plc.namePLC == item.namePLC).First().dataFC.Where(block => block.name == data.nameFC).First().code
//                    .AppendLine("REGION " + data.nameEq)
//                    .AppendLine("//Call functional for - " + data.nameEq)
//                    .AppendLine("\"" + data.name + "\"" + io.ToString())
//                    .AppendLine("END_REGION");
//            }
//            foreach (dataFunction dateFC in item.dataFC)
//            {
//                string blockString = dateFC.code.AppendLine("END_FUNCTION").ToString();
//                _tiaprj.CreateFC(item.namePLC, dateFC.name, dateFC.number, blockString, sourcePath, groupName:dateFC.group );
//            }
//            _tiaprj.Compile(item.namePLC);
//        }
//        _tiaprj.SaveProject();

//        if (closeProject) { _tiaprj.CloseProject(); }
//        return true;
//    }
//    catch (Exception ex)
//    {
//         msg = "[General]" + ex.Message;
//        return false;
//    }
//}

//private bool CreateFCFromExcelCallAllBlocks(bool closeProject = false)
//{
//    StringBuilder io;
//    StringBuilder param;
//    try
//    {
//        foreach (dataPLC item in _excelprj.BlocksStruct)
//        {                    
//            List<dataBlock> orderDataExcelListBlocks = item.instanceDB.OrderBy(order => order.name).ToList();                    
//            foreach (dataBlock data in orderDataExcelListBlocks)
//            {
//                io = null;
//                io = new StringBuilder();
//                param = null;
//                param = new StringBuilder();

//                if (data.typeEq.Length > 0 && data.nameEq.Length > 0)
//                {
//                    dataEq EqType = _excelprj.BlocksStruct
//                        .Where(plc => plc.namePLC == item.namePLC)
//                        .First()
//                        .Equipment
//                        .Where(eq => eq.typeEq == data.typeEq)
//                        .First();

//                    //Fill all needed symbol IO signals for block


//                    io.AppendLine("(");
//                    int numEqTag = EqType.dataTag.Count;
//                    foreach (dataTag tag in EqType.dataTag)
//                    {
//                        if (numEqTag == 1)
//                        {
//                            io.Append(tag.io + ":=" + data.nameEq + tag.name);
//                        }
//                        else
//                        {
//                            io.AppendLine(tag.io + ":=" + data.nameEq + tag.name + ",");
//                        }
//                        numEqTag--;
//                    }
//                    io.Append(");");

//                    //Fill all needed parameters for block

//                    param = new StringBuilder();
//                    if (EqType.dataParameter.Count > 1)
//                    {
//                        param.AppendLine("//Parameters for block: " + data.nameEq);
//                        foreach (dataParameter parameter in EqType.dataParameter)
//                        {
//                            if (parameter.type == "Input")
//                            {
//                                param.AppendLine("\"" + data.name + "\"." + parameter.io + ":=" + _excelprj.ModifyString(parameter.link, data.excelData) + ";");
//                            }
//                            if (parameter.type == "Output")
//                            {
//                                param.AppendLine(_excelprj.ModifyString(parameter.link, data.excelData) + ":=" + "\"" + data.name + "\"." + parameter.io + ";");
//                            }
//                        }
//                    }
//                }
//                if (!_excelprj.BlocksStruct.Where(plc => plc.namePLC == item.namePLC).First().dataFC.Any(block => block.name == data.nameFC))
//                {
//                    _excelprj.BlocksStruct.Where(plc => plc.namePLC == item.namePLC).First().dataFC.Add(new dataFunction()
//                    {
//                        name = data.nameFC,
//                        number = data.numFC,
//                        group = data.groupFC,
//                        code = new StringBuilder("FUNCTION " + "\"" + data.nameFC + "\" : Void\r\n")
//                                    .AppendLine("{ S7_Optimized_Access := 'TRUE' }")
//                                    .AppendLine("AUTHOR : SE")
//                                    .AppendLine("FAMILY : Constructor")
//                                    .AppendLine("BEGIN")
//                    });
//                }
//                _excelprj.BlocksStruct.Where(plc => plc.namePLC == item.namePLC).First().dataFC.Where(block => block.name == data.nameFC).First().code
//                .AppendLine("REGION " + data.nameEq)
//                .AppendLine("//Call functional block - " + data.instanceOfName + "for: " + data.nameEq)
//                .AppendLine("\"" + data.name + "\"" + io.ToString())
//                .AppendLine(param.ToString())
//                .AppendLine("END_REGION");

//            }
//            foreach (dataFunction dateFC in item.dataFC)
//            {
//                string blockString = dateFC.code.AppendLine("END_FUNCTION").ToString();
//                _tiaprj.CreateFC(item.namePLC, dateFC.name, dateFC.number, blockString, sourcePath, groupName: dateFC.group);
//            }
//            _tiaprj.Compile(item.namePLC);
//        }
//        _tiaprj.SaveProject();

//        if (closeProject) { _tiaprj.CloseProject(); }
//        return true;
//    }
//    catch (Exception ex)
//    {
//         msg = "[General]" + ex.Message;
//        return false;
//    }
//}
//}
//using (SEAppContext context = new SEAppContext())
//{
//    EqupmentList Equipment = new EqupmentList();
//    Equipment.Status = "Ok";
//    Equipment.EqNum = 710;
//    Equipment.EqName = "57_1";
//    Equipment.objName = "710";
//    Equipment.ObjType = "@_D";
//    Equipment.PicType = "1";
//    Equipment.ObjTagName = "710";
//    Equipment.PLCName = "PLC_1";
//    Equipment.InstanceOfName = "FB_Motor";
//    Equipment.Number_DB = 100;
//    Equipment.GroupDB = "@Eq_InstDB";
//    Equipment.GroupFC = "@Eq_FC";
//    Equipment.NameFC = "FC_EqCall";
//    Equipment.Number_FC = 100;

//    context.EqupmentLists.Add(Equipment);
//    context.SaveChanges();
//private bool CreateFCCallAllBlocsOld(string ProjectPath, bool closeProject = true)
//{
//string filename;
//string blockString;
//try
//{
//    filename = new FileInfo(ProjectPath).FullName;
//    _tiaprj.OpenProject(filename);
//    _tiaprj.ConnectTIA();
//    foreach (datePLC item in _excelprj.BlocksStruct)
//    {
//        foreach (dataBlock blockDB in item.instanceDB)
//        {
//            if (item.structFC.Any(name => name.name == blockDB.nameFC))
//                {
//                item.structFC.Where(name => name.name == blockDB.nameFC).First().code                            
//                    .AppendLine("REGION " + item.name.Remove(blockDB.nameDB.Length - 3))
//                    .AppendLine("//Call functional for - " + blockDB.nameDB.Remove(blockDB.nameDB.Length - 3))
//                    .AppendLine("\"" + blockDB.nameDB + "\"" + "();")
//                    .AppendLine("END_REGION");
//            }
//                else
//            {                            
//                item.structFC.Add(new dataFunction()
//                {
//                    name = blockDB.nameFC,
//                    number = blockDB.numberFC,
//                    group = blockDB.groupFC,
//                    code = new StringBuilder("FUNCTION " + "\"" + blockDB.nameFC + "\" : Void\r\n")
//                    .AppendLine("{ S7_Optimized_Access := 'TRUE' }")
//                    .AppendLine("AUTHOR : SE")
//                    .AppendLine("FAMILY : Constructor")
//                    .AppendLine("BEGIN")
//                    .AppendLine("REGION " + blockDB.nameDB.Remove(blockDB.nameDB.Length - 3))
//                    .AppendLine("//Call functional for - " + blockDB.nameDB.Remove(blockDB.nameDB.Length - 3))
//                    .AppendLine("\"" + blockDB.nameDB + "\"" + "();")
//                    .AppendLine("END_REGION"),
//            });
//            }
//        }
//        foreach (dateFC date in item.structFC)
//        {                        
//            blockString = date.code.AppendLine("END_FUNCTION").ToString();
//            _tiaprj.CreateFC(item.namePLC, date.name, date.number, blockString, sourcePath, date.group);
//        }
//        _tiaprj.Compile(item.namePLC);
//    }
//    _tiaprj.SaveProject();

//    if (closeProject) { _tiaprj.CloseProject(); }
//    return true;
//}
//catch (Exception ex)
//{
//     msg = "[General]" + ex.Message;
//    return false;
//}
//}


//private void UpdateMessage(string color = "Black")
//{
//    while (true)
//    {
//        if (msg != _tiaprj.Message)
//        {
//            _dispatcher.messageText(_tiaprj.Message);

//             msg = "[General]" + _tiaprj.Message;
//        }


//    }
//}