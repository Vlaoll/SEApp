using Microsoft.Win32;
using Siemens.Engineering.Library;
using Siemens.Engineering.SW.Blocks;
using Siemens.Engineering.SW.Tags;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Shapes;
using Excel = Microsoft.Office.Interop.Excel;
using seConfSW.Domain.Models;
using Microsoft.Extensions.Configuration;

namespace seConfSW.Views
{
    /// <summary>
    /// Interaction logic for HomeView.xaml
    /// </summary>
    public partial class HomeView : UserControl
    {
        private readonly IConfiguration configuration;
        public string Title { get; } = "Home";
        public static ThreadUI _dispatcher = new ThreadUI();
        private TIAPrj _tiaprj;
        private ExcelDataReader _excelprj;
        private string msg = string.Empty;
        private string sourceDBPath = @"Samples\sourceDB\";
        private string exportPath = @"Samples\export\";
        private string sourcePath = @"samples\source\";
        private string templatePath = @"samples\template\";
        //private string projectPath = @"TIA_main\TIA_main.ap18";
        private string projectPath = null;
        //private string projectLibPath = @"Library\Library.al18";
        private string projectLibPath = null;
        private string sourceTagPath = @"samples\tag\";
        //private string excelPath = @"Constructor.xlsx";
        private string excelPath = null;
        private string logPath = @"log\";
        private Siemens.Engineering.SW.PlcSoftware plcSoftware = null;
        FileStream Log = null;
        DateTime startTime;
        public HomeView()
        {
            InitializeComponent();
            if (DateTime.Now > new DateTime(2025, 06, 30, 23, 59, 59))
            {
                btnCheckPermission.IsEnabled = false;
            }
            else
            {
                btnCheckPermission.IsEnabled = true;
            }
            this.configuration = App.Configuration;
        }
        private bool CreateExcelDB()
        {
            
            try
            {
                DateTime temp = DateTime.Now;
                msg = temp + " :[General]Start to genetate excel DB";
                Trace.WriteLine(msg);

                _excelprj = null;
                _excelprj = new ExcelDataReader();
                msg = "[General]" + "##############################################################################################";
                Trace.WriteLine(msg);
                if (string.IsNullOrEmpty(excelPath))
                {
                    var filter = configuration["Excel:Filter"] ?? "Excel |*.xlsx;*.xlsm";
                    excelPath = _excelprj.SearchProject(filter: filter);
                }

                var mainSheetName = configuration["Excel:MainSheetName"] ?? "Main";
                _excelprj.OpenExcelFile(excelPath, mainSheetName: mainSheetName);
               

                if (!_excelprj.ReadExcelObjectData("Block",250))
                {
                    temp = DateTime.Now;
                    msg = temp + " :[General:Error] Wrong settings in excel file";
                    Trace.WriteLine(msg);
                    return false;
                }
                if (!_excelprj.ReadExcelExtendedData())
                {
                    temp = DateTime.Now;
                    msg = temp + " :[General:Error] Wrong settings in excel file";
                    Trace.WriteLine(msg);
                    return false;
                }
                _excelprj.CloseExcelFile();
                
                temp = DateTime.Now;
                msg = temp + " :[General]Finished to genetate excel DB";
                Trace.WriteLine(msg);
                msg = "[General]" + "##############################################################################################";
                Trace.WriteLine(msg);

                return true;
            }
            catch (Exception ex)
            {
                msg = "[General]" + ex.Message;
                Trace.WriteLine(msg);
                return false;
            }
        }

        private void StartTrace()
        {
            if (!Directory.Exists(logPath)) Directory.CreateDirectory(logPath);
            string LogPath = new FileInfo(logPath).FullName;
            string lodTime = DateTime.Now.ToString("yyyyMMdd'_'HHmm");
            Log = new FileStream(LogPath + "Log_" + lodTime + ".txt", FileMode.OpenOrCreate);
            Trace.Listeners.Add(new TextWriterTraceListener(Log));
            Trace.AutoFlush = true;
            Trace.WriteLine("Start log file: " + lodTime);
        }

        private bool InitTIAPrj(string ProjectPath)
        {
            if (Directory.Exists(exportPath)) Directory.Delete(exportPath, true);
            if (Directory.Exists(sourcePath)) Directory.Delete(sourcePath, true);
            if (Directory.Exists(templatePath)) Directory.Delete(templatePath, true);

            Directory.CreateDirectory(exportPath);
            Directory.CreateDirectory(sourcePath);
            Directory.CreateDirectory(templatePath);

            string filename;
            try
            {
                DateTime temp = DateTime.Now;
                msg = temp + " :[General]" + "Start to initialization TIA project";
                Trace.WriteLine(msg);

                _tiaprj.StartTIA();
                filename = new FileInfo(ProjectPath).FullName;
                _tiaprj.OpenProject(filename);
                _tiaprj.ConnectTIA();

                temp = DateTime.Now;
                msg = temp + " :[General]" + "Finished to connect to project";
                return true;
            }
            catch (Exception ex)
            {
                msg = "[General]" + ex.Message;
                Trace.WriteLine(msg);
                return false;
            }
        }
        private bool CteateTagTableFromFile(Siemens.Engineering.SW.PlcSoftware plcSoftware, dataPLC dataPLC, string groupName = "@Eq_TagTables", bool closeProject = false, bool saveProject = false, bool compileProject = false)
        {
            try
            {
                DateTime temp = DateTime.Now;
                msg = temp + " :[General]" + "Start creating tags";
                Trace.WriteLine(msg);

                List<dataImportTag> sources = new List<dataImportTag>();
                List<dataExcistTag> existTags = new List<dataExcistTag>(); ;

                if (Directory.Exists(sourceTagPath)) Directory.Delete(sourceTagPath, true);
                Directory.CreateDirectory(sourceTagPath);
                string filename;
                foreach (dataBlock instance in dataPLC.instanceDB)
                {
                    var tags = dataPLC.Equipment.Where(item => item.typeEq == instance.typeEq).First().dataTag;
                    foreach (dataTag tag in tags)
                    {

                        if (tag.adress != "")
                        {
                            filename = new FileInfo(sourceTagPath + tag.table + ".xml").FullName;
                            if (!sources.Any(s => s.name == tag.table))
                            {
                                sources.Add(new dataImportTag()
                                {
                                    path = filename,
                                    ID = 1,
                                    name = tag.table,
                                    code = new StringBuilder()
                                    .AppendLine("<?xml version=\"1.0\" encoding=\"utf-8\"?>\r\n<Document>")
                                    .AppendLine("\t<Engineering version=\"V18\" />  ")
                                    .AppendLine("\t<SW.Tags.PlcTagTable ID=\"" + 0 + "\">")
                                    .AppendLine("\t\t<AttributeList>")
                                    .AppendLine("\t\t\t<Name>" + tag.table + "</Name>")
                                    .AppendLine("\t\t</AttributeList>")
                                    .AppendLine("\t\t<ObjectList>")
                                });
                            }
                            string address = string.Empty;
                            var source = sources.Where(s => s.name == tag.table).First();
                            var link = Common.ModifyString(tag.link, instance.excelData);
                            (var exTagData, var isExist) = _tiaprj.FindTag(plcSoftware, link, tag.table);

                            if (isExist)
                            {
                                int ID = source.ID;
                                sources.Where(s => s.name == tag.table).First()
                                    .code.AppendLine("\t\t\t<SW.Tags.PlcTag ID=\"" + ID++ + "\" CompositionName=\"Tags\">")
                                    .AppendLine("\t\t\t\t<AttributeList>")
                                    .AppendLine("\t\t\t\t\t<DataTypeName>" + exTagData.type + "</DataTypeName>")
                                    .AppendLine("\t\t\t\t\t<LogicalAddress>%" + exTagData.adress + "</LogicalAddress>")
                                    .AppendLine("\t\t\t\t\t<Name>" + exTagData.link + "</Name>")
                                    .AppendLine("\t\t\t\t</AttributeList>")
                                    .AppendLine("\t\t\t\t<ObjectList>")
                                    .AppendLine("\t\t\t\t\t<MultilingualText ID=\"" + ID++ + "\" CompositionName=\"Comment\">")
                                    .AppendLine("\t\t\t\t\t\t<ObjectList>")
                                    .AppendLine("\t\t\t\t\t\t\t<MultilingualTextItem ID=\"" + ID++ + "\" CompositionName=\"Items\">")
                                    .AppendLine("\t\t\t\t\t\t\t\t<AttributeList>")
                                    .AppendLine("\t\t\t\t\t\t\t\t\t<Culture>en-US</Culture>")
                                    .AppendLine("\t\t\t\t\t\t\t\t\t<Text>" + exTagData.comment + "</Text>")
                                    .AppendLine("\t\t\t\t\t\t\t\t</AttributeList>")
                                    .AppendLine("\t\t\t\t\t\t\t</MultilingualTextItem>")
                                    .AppendLine("\t\t\t\t\t\t</ObjectList>")
                                    .AppendLine("\t\t\t\t\t</MultilingualText>")
                                    .AppendLine("\t\t\t\t</ObjectList>")
                                    .AppendLine("\t\t\t</SW.Tags.PlcTag>\"");
                                var newSources = sources.Where(s => s.name == tag.table).First();
                                newSources.ID = ID;
                                int index = sources.FindIndex(s => s.name == tag.table);
                                sources[index] = newSources;
                            }
                            else
                            {
                                var comments = Common.ModifyString(tag.comment, instance.excelData);
                                int ID = source.ID;
                                sources.Where(s => s.name == tag.table).First()
                                    .code.AppendLine("\t\t\t<SW.Tags.PlcTag ID=\"" + ID++ + "\" CompositionName=\"Tags\">")
                                    .AppendLine("\t\t\t\t<AttributeList>")
                                    .AppendLine("\t\t\t\t\t<DataTypeName>" + tag.type + "</DataTypeName>")
                                    .AppendLine("\t\t\t\t\t<LogicalAddress>%" + tag.adress + "</LogicalAddress>")
                                    .AppendLine("\t\t\t\t\t<Name>" + link + "</Name>")
                                    .AppendLine("\t\t\t\t</AttributeList>")
                                    .AppendLine("\t\t\t\t<ObjectList>")
                                    .AppendLine("\t\t\t\t\t<MultilingualText ID=\"" + ID++ + "\" CompositionName=\"Comment\">")
                                    .AppendLine("\t\t\t\t\t\t<ObjectList>")
                                    .AppendLine("\t\t\t\t\t\t\t<MultilingualTextItem ID=\"" + ID++ + "\" CompositionName=\"Items\">")
                                    .AppendLine("\t\t\t\t\t\t\t\t<AttributeList>")
                                    .AppendLine("\t\t\t\t\t\t\t\t\t<Culture>en-US</Culture>")
                                    .AppendLine("\t\t\t\t\t\t\t\t\t<Text>" + comments + "</Text>")
                                    .AppendLine("\t\t\t\t\t\t\t\t</AttributeList>")
                                    .AppendLine("\t\t\t\t\t\t\t</MultilingualTextItem>")
                                    .AppendLine("\t\t\t\t\t\t</ObjectList>")
                                    .AppendLine("\t\t\t\t\t</MultilingualText>")
                                    .AppendLine("\t\t\t\t</ObjectList>")
                                    .AppendLine("\t\t\t</SW.Tags.PlcTag>\"");
                                var newSources = sources.Where(s => s.name == tag.table).First();
                                newSources.ID = ID;
                                int index = sources.FindIndex(s => s.name == tag.table);
                                sources[index] = newSources;
                            }
                        }
                    }
                }
                foreach (var source in sources)
                {
                    source.code.AppendLine("\t\t</ObjectList>")
                        .AppendLine("\t</SW.Tags.PlcTagTable>\r\n</Document>");

                    // Write new file                        
                    using (var sw = new StreamWriter(source.path, true))
                    {
                        sw.WriteLine(source.code.ToString());
                    }
                    _tiaprj.ImportTagTable(plcSoftware, source.path);
                }
                temp = DateTime.Now;
                msg = temp + " :[General]" + "Finished creating tags";
                Trace.WriteLine(msg);
                if (compileProject) { _tiaprj.Compile(dataPLC.namePLC); ; }
                if (saveProject) { _tiaprj.SaveProject(); }
                if (closeProject) { _tiaprj.CloseProject(); }
                return true;
            }
            catch (Exception ex)
            {
                msg = "[General]" + ex.Message;
                Trace.WriteLine(msg);
                return false;
            }
        }
        private bool CteateTagsFromFile(Siemens.Engineering.SW.PlcSoftware plcSoftware, dataPLC dataPLC, string groupName = "@Eq_TagTables", bool closeProject = false, bool saveProject = false, bool compileProject = false)
        {
            try
            {
                DateTime temp = DateTime.Now;
                msg = temp + " :[General]" + "Start creating tags";
                Trace.WriteLine(msg);

                List<dataImportTag> sources = new List<dataImportTag>();
                List<dataExcistTag> existTags = new List<dataExcistTag>(); ;

                if (Directory.Exists(sourceTagPath)) Directory.Delete(sourceTagPath, true);
                Directory.CreateDirectory(sourceTagPath);
                string filename;

                PlcTagTableSystemGroup systemGroup = plcSoftware.TagTableGroup;
                PlcTagTableUserGroupComposition groupComposition = systemGroup.Groups;
                PlcTagTableUserGroup group = groupComposition.Find(groupName);

                foreach (dataBlock instance in dataPLC.instanceDB)
                {
                    var tags = dataPLC.Equipment.Where(item => item.typeEq == instance.typeEq).First().dataTag;
                    foreach (dataTag tag in tags)
                    {
                        if (tag.variant.Count == 0 || instance.variant.Intersect(tag.variant).Any())
                        {
                            if (tag.adress != "")
                            {
                                var link = Common.ModifyString(tag.link, instance.excelData);
                                bool isExist = _tiaprj.FindTag(plcSoftware, link, tag.table).isExist;
                                filename = new FileInfo(sourceTagPath + tag.table + ".xml").FullName;
                                if (!isExist)
                                {
                                    if (!sources.Any(s => s.name == tag.table))
                                    {
                                        sources.Add(new dataImportTag()
                                        {
                                            path = filename,
                                            table = tag.table,
                                            ID = 0,
                                            name = tag.table,
                                            code = new StringBuilder()
                                            .AppendLine("<?xml version=\"1.0\" encoding=\"utf-8\"?>\r\n<Document>")
                                            .AppendLine("\t<Engineering version=\"V18\" />  ")
                                        });
                                    }
                                    var source = sources.Where(s => s.name == tag.table).First();
                                    var comments = Common.ModifyString(tag.comment, instance.excelData);
                                    int ID = source.ID;
                                    sources.Where(s => s.name == tag.table).First().code
                                        .AppendLine("\t\t\t<SW.Tags.PlcTag ID=\"" + ID++ + "\" CompositionName=\"Tags\">")
                                        .AppendLine("\t\t\t\t<AttributeList>")
                                        .AppendLine("\t\t\t\t\t<DataTypeName>" + tag.type + "</DataTypeName>")
                                        .AppendLine("\t\t\t\t\t<LogicalAddress>%" + tag.adress + "</LogicalAddress>")
                                        .AppendLine("\t\t\t\t\t<Name>" + link + "</Name>")
                                        .AppendLine("\t\t\t\t</AttributeList>")
                                        .AppendLine("\t\t\t\t<ObjectList>")
                                        .AppendLine("\t\t\t\t\t<MultilingualText ID=\"" + ID++ + "\" CompositionName=\"Comment\">")
                                        .AppendLine("\t\t\t\t\t\t<ObjectList>")
                                        .AppendLine("\t\t\t\t\t\t\t<MultilingualTextItem ID=\"" + ID++ + "\" CompositionName=\"Items\">")
                                        .AppendLine("\t\t\t\t\t\t\t\t<AttributeList>")
                                        .AppendLine("\t\t\t\t\t\t\t\t\t<Culture>en-US</Culture>")
                                        .AppendLine("\t\t\t\t\t\t\t\t\t<Text>" + comments + "</Text>")
                                        .AppendLine("\t\t\t\t\t\t\t\t</AttributeList>")
                                        .AppendLine("\t\t\t\t\t\t\t</MultilingualTextItem>")
                                        .AppendLine("\t\t\t\t\t\t</ObjectList>")
                                        .AppendLine("\t\t\t\t\t</MultilingualText>")
                                        .AppendLine("\t\t\t\t</ObjectList>")
                                        .AppendLine("\t\t\t</SW.Tags.PlcTag>\"");
                                    var newSources = sources.Where(s => s.name == tag.table).First();
                                    newSources.ID = ID;
                                    int index = sources.FindIndex(s => s.name == tag.table);
                                    sources[index] = newSources;
                                }
                            }
                        }
                    }
                }
                foreach (var source in sources)
                {
                    source.code.AppendLine("</Document>");

                    // Write new file                        
                    using (var sw = new StreamWriter(source.path, true))
                    {
                        sw.WriteLine(source.code.ToString());
                    }
                    _tiaprj.ImportTags(plcSoftware, source.path, source.table);
                }
                temp = DateTime.Now;
                msg = temp + " :[General]" + "Finished creating tags";
                Trace.WriteLine(msg);
                if (compileProject) { _tiaprj.Compile(dataPLC.namePLC); ; }
                if (saveProject) { _tiaprj.SaveProject(); }
                if (closeProject) { _tiaprj.CloseProject(); }
                return true;
            }
            catch (Exception ex)
            {
                msg = "[General]" + ex.Message;
                Trace.WriteLine(msg);
                return false;
            }
        }
        private bool CteateUserConstants(Siemens.Engineering.SW.PlcSoftware plcSoftware, dataPLC dataPLC, bool closeProject = false, bool saveProject = false, bool compileProject = false)
        {
            try
            {
                DateTime temp = DateTime.Now;
                msg = temp + " :[General]" + "Start to creating user constants";
                Trace.WriteLine(msg);

                _tiaprj.CreateUserConstant(plcSoftware, dataPLC.userConstant);

                temp = DateTime.Now;
                msg = temp + " :[General]" + "Finished to creating user constants";
                Trace.WriteLine(msg);

                if (compileProject) { _tiaprj.Compile(dataPLC.namePLC); ; }
                if (saveProject) { _tiaprj.SaveProject(); }
                if (closeProject) { _tiaprj.CloseProject(); }
                return true;
            }
            catch (Exception ex)
            {
                msg = "[General]" + ex.Message;
                Trace.WriteLine(msg);
                return false;
            }
        }

        private bool CreateTags(Siemens.Engineering.SW.PlcSoftware plcSoftware, dataPLC dataPLC, bool closeProject = false, bool saveProject = false, bool compileProject = false)
        {
            try
            {
                DateTime temp = DateTime.Now;
                msg = temp + " [General]" + "Start creating tags";
                Trace.WriteLine(msg);
                _tiaprj.CreateTag(plcSoftware, dataPLC);

                temp = DateTime.Now;
                msg = temp + " [General]" + "Finished creating tags";
                Trace.WriteLine(msg);             

                if (compileProject) { _tiaprj.Compile(dataPLC.namePLC); ; }
                if (saveProject) { _tiaprj.SaveProject(); }
                if (closeProject) { _tiaprj.CloseProject(); }
                return true;
            }
            catch (Exception ex)
            {
                msg = "[General]" + ex.Message;
                Trace.WriteLine(msg);
                return false;
            }
        }
        private bool CreateEqConstants(Siemens.Engineering.SW.PlcSoftware plcSoftware, dataPLC dataPLC, bool closeProject = false, bool saveProject = false, bool compileProject = false)
        {
            try
            {
                DateTime temp = DateTime.Now;
                msg = temp + " [General]" + "Start creating equipment constants";
                Trace.WriteLine(msg);
                
               
                    foreach (dataBlock blockDB in dataPLC.instanceDB)
                    {
                        List<userConstant> listOfConstants = dataPLC.Equipment.Where(type => type.typeEq == blockDB.typeEq).FirstOrDefault().dataConstant;
                        if (listOfConstants != null)

                        _tiaprj.CreateUserConstant(plcSoftware, listOfConstants, blockDB.excelData, blockDB.nameEq);
                    }
                    temp = DateTime.Now;
                    msg = temp + " [General]" + "Finished creating equipment constants";
                    Trace.WriteLine(msg);
                        

                if (compileProject) { _tiaprj.Compile(dataPLC.namePLC); ; }
                if (saveProject) { _tiaprj.SaveProject(); }
                if (closeProject) { _tiaprj.CloseProject(); }
                return true;
            }
            catch (Exception ex)
            {
                msg = "[General]" + ex.Message;
                Trace.WriteLine(msg);
                return false;
            }
        }

        private bool ConnectLib(Siemens.Engineering.SW.PlcSoftware plcSoftware, dataPLC dataPLC, string LibraryPath, bool closeProject = false, bool saveProject = false, bool compileProject = false)
        {
            string fileNameLibrary;
            try
            {
                msg = "[General]" + "Start updating project library";
                Trace.WriteLine(msg);
                fileNameLibrary = new FileInfo(LibraryPath).FullName;
                _tiaprj.UpdatePrjLibraryFromGlobal(fileNameLibrary);
                if (compileProject) { _tiaprj.Compile(dataPLC.namePLC); ; }
                if (saveProject) { _tiaprj.SaveProject(); }
                if (closeProject) { _tiaprj.CloseProject(); }
                return true;
            }
            catch (Exception ex)
            {
                msg = "[General]" + ex.Message;
                Trace.WriteLine(msg);
                return false;
            }
        }
        private bool UpdateSupportBlocks(Siemens.Engineering.SW.PlcSoftware plcSoftware, dataPLC dataPLC, string LibraryPath, bool closeProject = false, bool saveProject = false, bool compileProject = false)
        {
            string fileNameLibrary;
            try
            {
                msg = "[General]" + "Start updating support blocks from library";
                Trace.WriteLine(msg);
                fileNameLibrary = new FileInfo(LibraryPath).FullName;
             

                //Add needed DBs and UDTs blocks from MC in global library
                UserGlobalLibrary globalLibrary = _tiaprj.OpenLibrary(fileNameLibrary);
                foreach (dataSupportBD item in dataPLC.dataSupportBD)
                {
                    try
                    {
                        if (item.type == "DB" && !item.isType)
                        {
                            _tiaprj.CopyBlocksFromMasterCopyFolder(plcSoftware, globalLibrary, item.path + "." + item.name, item.group);
                        }
                        else if (item.type == "UDT" && !item.isType)
                        {
                            _tiaprj.CopyUDTFromMasterCopyFolder(plcSoftware, globalLibrary, item.path + "." + item.name, item.group);
                        }

                        //Add needed UDT from project library
                        else if (item.type == "UDT" && item.isType)
                        {
                            _tiaprj.GenerateUDTFromLibrary(plcSoftware, "Common", item.name, item.group);
                        }
                    }
                    catch (Exception ex)
                    {
                        msg = "[General]" + ex.Message + " :" + item.name;
                        Trace.WriteLine(msg);
                        return false;
                    }

                }                
                globalLibrary.Close();             

                if (compileProject) { _tiaprj.Compile(dataPLC.namePLC); ; }
                if (saveProject) { _tiaprj.SaveProject(); }
                if (closeProject) { _tiaprj.CloseProject(); }
                return true;
            }
            catch (Exception ex)
            {
                msg = "[General]" + ex.Message;
                Trace.WriteLine(msg);
                return false;
            }
        }
        private bool UpdateTypeBlocks(Siemens.Engineering.SW.PlcSoftware plcSoftware, dataPLC dataPLC, string LibraryPath, bool closeProject = false, bool saveProject = false, bool compileProject = false, bool ClearProjectLibrary = false)
        {
            string fileNameLibrary;
            try
            {
                msg = "[General]" + "Start updating types blocks from project library";
                Trace.WriteLine(msg);
                fileNameLibrary = new FileInfo(LibraryPath).FullName;
                UserGlobalLibrary globalLibrary = _tiaprj.OpenLibrary(fileNameLibrary);

                //Add needed functional block type from project library
                foreach (dataEq equipment in dataPLC.Equipment)
                {
                    foreach (dataLibrary block in equipment.FB)
                    {
                        try
                        {
                            if (!block.isType)
                            {
                                _tiaprj.CopyBlocksFromMasterCopyFolder(plcSoftware, globalLibrary, block.path + "." + block.name, block.group);
                            }
                            else
                            {
                                _tiaprj.GenerateBlockFromLibrary(plcSoftware, block.path, block.name, block.group);
                            }
                        }
                        catch (Exception ex)
                        {
                            msg = "[General]" + ex.Message + " :" + block.name;
                            Trace.WriteLine(msg);
                            return false;
                        }
                    }  
                }
                globalLibrary.Close();
                
                if (ClearProjectLibrary) { _tiaprj.CleanUpMainLibrary(); }
                if (compileProject) { _tiaprj.Compile(dataPLC.namePLC); }
                if (saveProject) { _tiaprj.SaveProject(); }
                if (closeProject) { _tiaprj.CloseProject(); }
                return true;
            }
            catch (Exception ex)
            {
                msg = "[General]" + ex.Message;
                Trace.WriteLine(msg);
                return false;
            }
        }
        private bool AddValueToDataBlock(Siemens.Engineering.SW.PlcSoftware plcSoftware, dataPLC dataPLC, bool closeProject = false, bool saveProject = false, bool compileProject = false)
        {
            List<dataDataBlockValue> dataDBValues = new List<dataDataBlockValue>();
            List<dataSupportBD> dataSupportBD = new List<dataSupportBD>() ;            
            Dictionary<string, StringBuilder> DictDBBase = new Dictionary<string, StringBuilder>();
            Dictionary < string,Dictionary<string,string>> DictDBValues = new Dictionary<string,Dictionary<string,string>>();
            StringBuilder codeSB;
            try
            {
                foreach (var instance in dataPLC.instanceDB)
                {

                    foreach (var equipment in dataPLC.Equipment.Where(type => type.typeEq == instance.typeEq).FirstOrDefault().dataDataBlockValue)
                    {
                        string name = Common.ModifyString(equipment.name, instance.excelData);
                        if (!DictDBValues.ContainsKey(equipment.DB))
                        {
                            DictDBValues.Add(equipment.DB, new Dictionary<string, string>());
                        }
                        if (!DictDBValues[equipment.DB].ContainsKey(name))
                        {
                            DictDBValues[equipment.DB].Add(name, " : " + equipment.type + ";");
                        }
                        else
                        {
                            DictDBValues[equipment.DB][name] = " : " + equipment.type + ";";
                        }
                    }
                   
                                      
                    foreach (var equipment in dataPLC.Equipment.Where(type => type.typeEq == instance.typeEq).FirstOrDefault().dataSupportBD)
                    {
                        if (!dataSupportBD.Any(item =>item.name == equipment.name))
                        {
                            dataSupportBD.Add(new dataSupportBD()
                            {
                                name = equipment.name,
                                number = equipment.number,
                                group = equipment.group,                                
                                path = equipment.path,
                                isType = equipment.isType,
                                isRetain = equipment.isRetain,
                                isOptimazed = equipment.isOptimazed,
                            });
                        }
                    }
                }
                foreach (var DB in dataSupportBD)
                {
                    codeSB = null;
                    codeSB = new StringBuilder();
                    var group = _tiaprj.CreateBlockGroup(plcSoftware, DB.group);
                    var filename = _tiaprj.GenerateSourceBlock(plcSoftware, DB.group + "." + DB.name, sourcePath);
                    if (!DictDBBase.ContainsKey(DB.name) && filename == string.Empty)
                    {
                        codeSB.AppendLine("DATA_BLOCK \"" + DB.name + "\"");
                        if (DB.isOptimazed)
                        {
                            codeSB.AppendLine("{ S7_Optimized_Access := 'TRUE' }");
                        }
                        else
                        {
                            codeSB.AppendLine("{ S7_Optimized_Access := 'FALSE' }");
                        }
                        codeSB.AppendLine("VERSION : 0.1");
                        if (!DB.isRetain)
                        {
                            codeSB.AppendLine("NON_RETAIN");
                        }
                        codeSB.AppendLine("STRUCT")
                            .AppendLine("END_STRUCT;")
                            .AppendLine("BEGIN")
                            .AppendLine("END_DATA_BLOCK");
                        DictDBBase.Add(DB.name, codeSB);

                        using (var sw = new StreamWriter(sourcePath + DB.name + ".db"))
                        {
                            foreach (var item in codeSB.ToString().Replace('\r', ' ').Split('\n'))
                            {
                                sw.WriteLine(item);
                            }  
                            filename = sourcePath + DB.name + ".db";
                        }
                        _tiaprj.ClearSource(plcSoftware);
                        _tiaprj.ImportSource(plcSoftware, DB.name, filename);
                        _tiaprj.ChangeBlockNumber(plcSoftware, DB.name, DB.number, _tiaprj.GenerateBlock(plcSoftware, DB.group));
                        msg = "[General]" + "Created empty source for: " + DB.name;
                        Trace.WriteLine(msg);
                    }
                    if (filename != string.Empty)
                    {
                        var tempFilename = System.IO.Path.GetTempFileName();
                        // Read file
                        using (var sr = new StreamReader(filename))
                        {
                            // Write new file
                            using (var sw = new StreamWriter(tempFilename))
                            {
                                // Read lines
                                string line;                                
                                while ((line = sr.ReadLine()) != null)
                                {                                         
                                    if (line.Contains("STRUCT") || line.Contains("VAR RETAIN") )
                                    {
                                        string[] str;
                                        line = sr.ReadLine();
                                        while (line != null && (!line.Contains("END_VAR") && !line.Contains("END_STRUCT;")))
                                        {
                                            str = Array.Empty<string>();
                                            str = line.Replace(" ", "").Split(':');
                                            if (str.Length > 1 && str.Length<3)
                                            {
                                                if (!DictDBValues.ContainsKey(DB.name))
                                                {
                                                    DictDBValues.Add(DB.name, new Dictionary<string, string>());
                                                }
                                                if (!DictDBValues[DB.name].ContainsKey(str[0]))
                                                {
                                                    DictDBValues[DB.name].Add(str[0], " : " + str[1]);
                                                }
                                                else
                                                {
                                                    DictDBValues[DB.name][str[0]] = " : " + str[1];
                                                }
                                            }
                                            line = sr.ReadLine();
                                        }                                       
                                    }
                                    else if(line.Contains("BEGIN"))
                                    {
                                        break;
                                    }
                                    else 
                                    {
                                        // Keep lines that does not match
                                        sw.WriteLine(line);
                                    }                                    
                                }
                                if (DB.isRetain)
                                {
                                    //var dict = DictDBValues[DB.name].OrderBy(key => key.Key);
                                    sw.WriteLine("VAR RETAIN");
                                    foreach (var item in DictDBValues[DB.name])
                                    {
                                        sw.WriteLine(item.Key + item.Value);
                                    }
                                    sw.WriteLine("END_VAR");
                                    sw.WriteLine("BEGIN");
                                    sw.WriteLine("END_DATA_BLOCK");
                                }
                                else
                                {
                                    //var dict = DictDBValues[DB.name].OrderBy(key => key.Key);
                                    sw.WriteLine("STRUCT");
                                    foreach (var item in DictDBValues[DB.name])
                                    {
                                        sw.WriteLine(item.Key + item.Value);
                                    }
                                    sw.WriteLine("END_STRUCT;");
                                    sw.WriteLine("BEGIN");
                                    sw.WriteLine("END_DATA_BLOCK");
                                }    
                            }
                        }
                        // Delete original file
                        File.Delete(filename);

                        // ... and put the temp file in its place.
                        File.Move(tempFilename, filename);
                        File.Delete(tempFilename);

                        _tiaprj.ClearSource(plcSoftware);
                        _tiaprj.ImportSource(plcSoftware, DB.name, filename);
                        _tiaprj.ChangeBlockNumber(plcSoftware, DB.name, DB.number, _tiaprj.GenerateBlock(plcSoftware, DB.group));
                        msg = "[General]" + "Generated new  source for: " + DB.name;
                        Trace.WriteLine(msg);
                    }
                }
                
                return false;
            }
            catch (Exception)
            {

                return false;
            }
        }
        private bool CreateInstanceBlocks(Siemens.Engineering.SW.PlcSoftware plcSoftware, dataPLC dataPLC, bool closeProject = false, bool saveProject = false, bool compileProject = false)
        {
            try
            {
                DateTime temp = DateTime.Now;
                msg = temp + "[General]" + "Start creating instance blocks";
                Trace.WriteLine(msg);
                foreach (dataBlock blockDB in dataPLC.instanceDB)
                {
                    int numberDB = blockDB.number != 0 ?(blockDB.number-1) * 20 + 1000:0;
                    if (!dataPLC.Equipment.Where(type => type.typeEq == blockDB.typeEq).FirstOrDefault().isExtended)
                    {
                        _tiaprj.CreateInstanceDB(plcSoftware, blockDB.name, numberDB, blockDB.instanceOfName, sourceDBPath, blockDB.group, blockDB.typeEq);
                    }
                    else
                    {

                        foreach (dataExtSupportBlock item in dataPLC.Equipment.Where(type => type.typeEq == blockDB.typeEq).FirstOrDefault().dataExtSupportBlock)
                        {
                            if (item.variant.Count == 0 || blockDB.variant.Intersect(item.variant).Any())
                            {
                                string name = Common.ModifyString(item.name, blockDB.excelData);
                                _tiaprj.CreateInstanceDB(plcSoftware, name, numberDB + item.number, item.instanceOfName, sourceDBPath, blockDB.group, blockDB.typeEq);
                            }
                            
                        }
                    }
                }
                if (compileProject) { _tiaprj.Compile(dataPLC.namePLC); ; }
                if (saveProject) { _tiaprj.SaveProject(); }
                if (closeProject) { _tiaprj.CloseProject(); }
                return true;
            }
            catch (Exception ex)
            {
                msg = "[General]" + ex.Message;
                Trace.WriteLine(msg);
                return false;
            }
        }
        private bool EditFCFromExcelCallAllBlocks(Siemens.Engineering.SW.PlcSoftware plcSoftware, dataPLC dataPLC, bool closeProject = false, bool saveProject = false, bool compileProject = false)
        {
            StringBuilder io;
            StringBuilder param;
            string[] code;
            var sources = new Dictionary<string, string>();

            DateTime temp = DateTime.Now;
            msg = temp + "[General]" + "Start editing FCs";
            Trace.WriteLine(msg);

            try
            {
                sources.Clear();
                List<dataBlock> orderDataExcelListBlocks = dataPLC.instanceDB.Where(element => dataPLC.Equipment.Where(type => type.typeEq == element.typeEq).FirstOrDefault().isExtended == false)
                    .OrderBy(order => order.number)
                    .ToList();
                //List<dataBlock> orderDataExcelListBlocks = dataPLC.instanceDB.Where(type => type.typeEq != dataPLC.Equipment.Where(typeEq => typeEq.isExtended == true).FirstOrDefault().typeEq).OrderBy(order => order.nameEq).ToList();
                foreach (dataBlock data in orderDataExcelListBlocks)
                {
                    io = null;
                    io = new StringBuilder();
                    param = null;
                    param = new StringBuilder();
                    code = null;

                    if (data.typeEq.Length > 0 && data.nameEq.Length > 0)
                    {
                        dataEq EqType = _excelprj.BlocksStruct
                            .Where(plc => plc.namePLC == dataPLC.namePLC)
                            .First()
                            .Equipment
                            .Where(eq => eq.typeEq == data.typeEq)
                            .First();

                        //Fill all needed symbol IO signals for block
                        io.AppendLine("(");
                        int numEqTag = EqType.dataTag.Count;
                        foreach (dataTag tag in EqType.dataTag)
                        {
                            if (numEqTag == 1)
                            {          
                                io.Append(tag.name + ":=" + Common.ModifyString(tag.link, data.excelData));

                            }
                            else
                            {                                
                                io.AppendLine(tag.name + ":=" + Common.ModifyString(tag.link, data.excelData) + ",");
                            }
                            numEqTag--;
                        }
                        io.Append(");");

                        //Fill all needed parameters for block
                        param = new StringBuilder();
                        if (EqType.dataParameter.Count > 1)
                        {
                            param.AppendLine("//Parameters for block: " + data.nameEq);
                            foreach (dataParameter parameter in EqType.dataParameter)
                            {
                                if (parameter.type == "I") { param.AppendLine("\"" + data.name + "\"." + parameter.name + ":=" + Common.ModifyString(parameter.link, data.excelData) + ";"); }
                                if (parameter.type == "O") { param.AppendLine(Common.ModifyString(parameter.link, data.excelData) + ":=" + "\"" + data.name + "\"." + parameter.name + ";"); }
                            }
                        }
                    }
                    code = new StringBuilder()
                            .AppendLine("REGION " + "Call FB - \"" + data.instanceOfName + "\" for: " + data.nameEq)
                            .AppendLine("//Call functional block - " + data.instanceOfName + " for: " + data.nameEq)
                            .AppendLine("//" + data.comment)
                            .AppendLine("\"" + data.name + "\"" + io.ToString())
                            .AppendLine(param.ToString())
                            .AppendLine("END_REGION")
                            .AppendLine("END_FUNCTION")
                            .ToString()
                            .Replace('\r', ' ')
                            .Split('\n');

                    if (!sources.ContainsKey(data.nameFC))
                    {
                        dataFunction dateFC = dataPLC.dataFC.Where(name => name.name == data.nameFC).First();
                        (PlcBlockGroup group, string nameBlock) = _tiaprj.GetBlockGroup(plcSoftware, dateFC.group + "." + dateFC.name);

                        PlcBlock plcBlock = group != null ? group.Blocks.Find(nameBlock) : null;
                        if (plcBlock == null)
                        {
                            string blockString = dateFC.code.AppendLine("END_FUNCTION").ToString();
                            _tiaprj.CreateFC(plcSoftware, dateFC.name, dateFC.number, blockString, sourcePath, groupName: dateFC.group);
                        }

                        sources.Add(data.nameFC, _tiaprj.GenerateSourceBlock(plcSoftware, dateFC.group + "." + dateFC.name, sourcePath));
                    }
                    if (sources.Where(key => key.Key == data.nameFC).First().Value != "")
                    {
                        var path = DeleteBlockRegion(sources.Where(key => key.Key == data.nameFC).First().Value, data.nameEq);
                        AddBlockRegion(path, code, data.nameEq);
                    }
                }
                foreach (var source in sources)
                {
                    _tiaprj.CreateFC(plcSoftware, source.Key, dataPLC.dataFC.Where(fc => fc.name == source.Key).First().number, source.Value, dataPLC.dataFC.Where(fc => fc.name == source.Key).First().group);

                }

                _tiaprj.ClearSource(plcSoftware);
                if (compileProject) { _tiaprj.Compile(dataPLC.namePLC); ; }
                if (saveProject) { _tiaprj.SaveProject(); }
                if (closeProject) { _tiaprj.CloseProject(); }


                return true;
            }
            catch (Exception ex)
            {
                msg = "[General]" + ex.Message;
                Trace.WriteLine(msg);
                return false;
            }
        }
        private bool CreateTemplateFCFromExcel(Siemens.Engineering.SW.PlcSoftware plcSoftware, dataPLC dataPLC, bool closeProject = false, bool saveProject = false, bool compileProject = false)
        {
            var sources = new Dictionary<string, string>();
            var files = new Dictionary<string, string>();
            string[] code;
            if (Directory.Exists(templatePath)) Directory.Delete(templatePath, true);
            Directory.CreateDirectory(templatePath);

            DateTime temp = DateTime.Now;
            msg = temp + "[General]" + "Start tempales FCs";
            Trace.WriteLine(msg);

            try
            {
                List<dataBlock> orderDataExcelListBlocks = dataPLC.instanceDB.Where(element => dataPLC.Equipment.Where(type => type.typeEq == element.typeEq).FirstOrDefault().isExtended == true)
                    .OrderBy(order => order.number)
                    .ToList();
                
                //Export sources for all templates
                var Eqs = dataPLC.Equipment.Where(typeEq => typeEq.isExtended == true);
                foreach (dataEq type in Eqs)
                {
                    dataLibrary template = type.FB.FirstOrDefault();
                    if (template.GetType() != null)
                    {
                        (PlcBlockGroup group, string nameBlock) = _tiaprj.GetBlockGroup(plcSoftware, template.group + "." + template.name);

                        PlcBlock plcBlock = group != null ? group.Blocks.Find(nameBlock) : null;
                        if (plcBlock != null)
                        {
                            _tiaprj.GenerateSourceBlock(plcSoftware, template.group + "." + template.name, templatePath);

                        }
                    }
                }
                foreach (dataBlock data in orderDataExcelListBlocks)
                {
                    // Read file
                    string filename = new FileInfo(templatePath + data.instanceOfName + ".scl").FullName;
                    string eqFileName = new FileInfo(templatePath + data.group + "." + "FC_Call-" + data.typeEq + ".scl").FullName;
                    code = null;
                    if (!files.ContainsKey(data.group + "." + "FC_Call-" + data.typeEq))
                    {
                        files.Add(data.group + "." + "FC_Call-" + data.typeEq, eqFileName);
                    }

                    using (var sr = new StreamReader(filename))
                    {
                        // Write new file
                        using (var sw = new StreamWriter(eqFileName, true))
                        {
                            // Read lines
                            string line, lineNext;
                            line = sr.ReadLine();
                            if (line.Contains("FUNCTION \"" + data.instanceOfName) && line != null)
                            {
                                sw.WriteLine("FUNCTION \"" + data.name + "\" : Void");

                            }
                           
                            while ((line = sr.ReadLine()) != null)
                            {
                                if (line.Contains("REGION") && !line.Contains("END_REGION") && line != null)
                                {

                                    if ( (lineNext = sr.ReadLine()).Contains("isVariant") && lineNext != null)
                                    {
                                        var n = lineNext.IndexOf("isVariant")  + 12;
                                        List<string> str = lineNext.Remove(0, n).Split(',').ToList();
                                        var isVariantExist = data.variant.Intersect(str).ToList();
                                        if (!isVariantExist.Any())
                                        {
                                            while ((line = sr.ReadLine()) != "\tEND_REGION" && line!= null) { }

                                        }
                                        else
                                        {
                                            sw.WriteLine(Common.ModifyString(line, data.excelData,"@"));
                                            sw.WriteLine(lineNext);
                                        }
                                    }
                                    else
                                    {
                                        sw.WriteLine(Common.ModifyString(line, data.excelData, "@"));
                                        sw.WriteLine(Common.ModifyString(lineNext, data.excelData, "@"));
                                    }   
                                }
                                //else if (line.Contains(": Variant;")) sw.WriteLine(line, data.excelData);
                                else sw.WriteLine(Common.ModifyString(line, data.excelData, "@"));
                            }
                        }
                    }
                    code = new StringBuilder()                            
                            .AppendLine("REGION " + "Call template - \"" + data.instanceOfName + "\" for: " + data.nameEq)                            
                            .AppendLine("//" + data.comment)
                            .AppendLine(" \"" + data.name + "\"();")
                            .AppendLine("END_REGION")
                            .AppendLine("END_FUNCTION")
                            .ToString()
                            .Replace('\r', ' ')
                            .Split('\n');
                    if (!sources.ContainsKey(data.nameFC))
                    {
                        dataFunction dateFC = dataPLC.dataFC.Where(name => name.name == data.nameFC).First();
                        (PlcBlockGroup group, string nameBlock) = _tiaprj.GetBlockGroup(plcSoftware, dateFC.group + "." + dateFC.name);

                        PlcBlock plcBlock = group != null ? group.Blocks.Find(nameBlock) : null;
                        if (plcBlock == null)
                        {
                            string blockString = dateFC.code.AppendLine("END_FUNCTION").ToString();
                            _tiaprj.CreateFC(plcSoftware, dateFC.name, dateFC.number, blockString, sourcePath, groupName: dateFC.group);
                        }

                        sources.Add(data.nameFC, _tiaprj.GenerateSourceBlock(plcSoftware, dateFC.group + "." + dateFC.name, sourcePath));
                    }
                    if (sources.Where(key => key.Key == data.nameFC).First().Value != "")
                    {
                        var path = DeleteBlockRegion(sources.Where(key => key.Key == data.nameFC).First().Value, data.nameEq);
                        AddBlockRegion(path, code, data.nameEq);
                    }
                }
                foreach (var file in files)
                {
                    _tiaprj.CreateFC(plcSoftware, file.Key, 0, file.Value, file.Key);
                }
                foreach (var source in sources)
                {
                    _tiaprj.CreateFC(plcSoftware, source.Key, dataPLC.dataFC.Where(fc => fc.name == source.Key).First().number, source.Value, dataPLC.dataFC.Where(fc => fc.name == source.Key).First().group);
                }

                _tiaprj.ClearSource(plcSoftware);
                plcSoftware.BlockGroup.Groups.Find("@Template").Delete();

                if (compileProject) { _tiaprj.Compile(dataPLC.namePLC); ; }
                if (saveProject) { _tiaprj.SaveProject(); }
                if (closeProject) { _tiaprj.CloseProject(); }


                return true;
            }
            catch (Exception ex)
            {
                msg = "[General]" + ex.Message;
                Trace.WriteLine(msg);
                return false;
            }
        }


        private string DeleteBlockRegion(string filename, string eqName)
        {
            var tempFilename = System.IO.Path.GetTempFileName();
            try
            {
                // Read file
                using (var sr = new StreamReader(filename))
                {
                    // Write new file
                    using (var sw = new StreamWriter(tempFilename))
                    {
                        // Read lines
                        string line;
                        bool exit = false;
                        while ((line = sr.ReadLine()) != null)
                        {
                            // Look for text to remove
                            int numLine = line.Trim().Length;
                            //string str = "REGION " + eqName;
                            if (line.Contains("REGION Call") && line.Contains(eqName))//&& (numLine == str.Length))
                            {
                                exit = true;
                                msg = "[General]" + "Delete Region for block: " + eqName;
                                Trace.WriteLine(msg);
                            }
                            if (!line.Contains("REGION " + eqName) && !exit)
                            {
                                // Keep lines that does not match
                                sw.WriteLine(line);
                            }
                            else
                            {
                                // Search END REGION 
                                exit = line.Contains("END_REGION") ? false : true;
                            }
                        }
                    }
                }
                // Delete original file
                File.Delete(filename);

                // ... and put the temp file in its place.
                File.Move(tempFilename, filename);
                File.Delete(tempFilename);
            }
            catch (Exception e)
            {
                msg = "[General]" + "Exception: " + e.Message;
            }
            return filename;
        }
        private bool AddBlockRegion(string filename, string[] newBlock, string eqName)
        {
            var tempFilename = System.IO.Path.GetTempFileName();
            // Read file
            using (var sr = new StreamReader(filename))
            {
                // Write new file
                using (var sw = new StreamWriter(tempFilename))
                {
                    // Read lines
                    string line;

                    while ((line = sr.ReadLine()) != null)
                    {
                        // Look for text to remove
                        if (!line.Contains("END_FUNCTION"))
                        {
                            // Keep lines that does not match
                            sw.WriteLine(line);
                        }
                        else
                        {
                            foreach (var item in newBlock)
                            {
                                sw.WriteLine(item);
                            }
                            // Add  lines for Region with new block 
                            msg = "[General]" + "Add Region for block: " + eqName;
                            Trace.WriteLine(msg);
                        }
                    }
                }
            }
            // Delete original file
            File.Delete(filename);

            // ... and put the temp file in its place.
            File.Move(tempFilename, filename);
            File.Delete(tempFilename);
            return true;
        }     
        private void BtnCheckPermission_Click(object sender, RoutedEventArgs e)
        {
            //Check installed version of Excel           
            
            try
            {
                var oXL = new Excel.Application();                
                int versionExel = Convert.ToInt32(oXL.Version.Substring(0, 2));
                oXL.Quit();
                if (versionExel < 16)
                {
                    msg = "Incorrect version for Excel. Less then 2016 " ;
                    tbMessage.Text = msg;
                    return;

                }
            }
            catch (Exception)
            {
                //msg = ex.Message + " " + v + " " + t + " " + x;
                msg = "Missing installed Excel or Incorrect version for Excel. Less then 2016";
                tbMessage.Text = msg;
                return;
            }

            // check is user in group
            System.Security.Principal.WindowsPrincipal principal = new System.Security.Principal.WindowsPrincipal(System.Security.Principal.WindowsIdentity.GetCurrent());
            if (!principal.IsInRole("Siemens TIA Openness"))
            {
                msg = "Add user to group: Siemens TIA Openness";
                tbMessage.Text = msg;
                return;
            }

            //check TIA Portal version
            RegistryKey filePathReg = Registry.LocalMachine.OpenSubKey("SOFTWARE\\Siemens\\Automation\\Openness\\19.0\\PublicAPI\\19.0.0.0");
            if (filePathReg == null)
            {
                msg = "Missing or Incorrect version ofTIA Portal. Must be TIA Portal v19";
                tbMessage.Text = msg;
                return;
            }
            msg = "All permissions are ok";
            tbMessage.Text = msg;
            btnReadExcel.IsEnabled = true;
            btnCheckPermission.IsEnabled = false;
        }
        private void BtnReadExcel_Click(object sender, RoutedEventArgs e)
        {
            btnCheckPermission.IsEnabled = false;
            startTime = DateTime.Now;
            StartTrace();
            tbMessage.Text = "Start to read Excel file";
            bool isNext = CreateExcelDB();
            if (isNext)
            {
                tbMessage.Text = "Reading Excel file is completed. Select needed action for TIA project";
                btnConnectToOpened.IsEnabled = true;
                btnOpenTIA.IsEnabled = true;
                btnReadExcel.IsEnabled = false;

                chbVisible.IsEnabled = true;
                chbClose.IsEnabled = true;
                chbCompile.IsEnabled = true;
                chbSave.IsEnabled = true;
                return;
            }
            tbMessage.Text = "Reading Excel file is wrong.";
        }
        private void BtnConnectToOpened_Click(object sender, RoutedEventArgs e)
        {

            chbVisible.IsChecked = true;
            chbVisible.IsEnabled = false;
            bool isVible = chbVisible.IsChecked == false || chbVisible.IsChecked == null ? false : true;
            _tiaprj = null;
            _tiaprj = new TIAPrj(isVible);


            tbMessage.Text = "Try connect to opened TIA project";
            bool isNext = _tiaprj.ConnectTIA() == 2 ? true : false;
            if (isNext)
            {
                tbMessage.Text = "Connecting is successful. Pls select path for library";
                btnReadExcel.IsEnabled = false;
                btnOpenTIA.IsEnabled = false;
                btnConnectToOpened.IsEnabled = false;
                btnSelectLibrary.IsEnabled = true;
                btnExecute.IsEnabled = false;

                chbCreateTags.IsEnabled = true;
                chbCreateInsDB.IsEnabled = true;
                chbCreateFC.IsEnabled = true;

                chbClose.IsEnabled = false;
                chbClose.IsChecked = false;
                chbCompile.IsEnabled = true;
                chbSave.IsEnabled = true;
                return;
            }
            tbMessage.Text = "Connecting to TIA project was wrong. Try again or open project";
        }

        
        private void BtnOpenTIA_Click(object sender, RoutedEventArgs e)
        {
            bool isVible = chbVisible.IsChecked == false || chbVisible.IsChecked == null ? false : true;
            _tiaprj = null;
            _tiaprj = new TIAPrj(isVible);

            tbMessage.Text = "Try  to open new TIA project";
            if (_tiaprj.ConnectTIA() == 0)
            {
                if (string.IsNullOrEmpty(projectPath))
                {
                    projectPath = _tiaprj.SearchProject();
                }
                
                if (!string.IsNullOrEmpty(projectPath))
                {
                    if (_tiaprj.ConnectTIA() == 2)
                    {
                        btnReadExcel.IsEnabled = false;
                        btnConnectToOpened.IsEnabled = false;
                        btnOpenTIA.IsEnabled = false;
                        btnSelectLibrary.IsEnabled = true;
                        btnExecute.IsEnabled = false;                       
                        chbCreateTags.IsEnabled = true;
                        chbCreateInsDB.IsEnabled = true;
                        chbCreateFC.IsEnabled = true;
                        if (!isVible)
                        {
                            chbClose.IsEnabled = false;
                            chbClose.IsChecked = true;
                            chbCompile.IsEnabled = false;
                            chbCompile.IsChecked = true;
                            chbSave.IsEnabled = false;
                            chbSave.IsChecked = true;
                        }
                        else
                        {
                            chbClose.IsEnabled = true;                            
                            chbCompile.IsEnabled = true;                           
                            chbSave.IsEnabled = true;                           
                        }
                        tbMessage.Text = "TIA project is opened and connected. Pls select path for library";
                        return;
                    }
                }
                tbMessage.Text = "TIA didn't selected or wrong path";
                return;
            }
            tbMessage.Text = "Please close TIA Portal.";
        }
        private void BtnSelectLibrary_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(projectLibPath))
            {
                projectLibPath = _tiaprj.SearchLibrary();
            }

            if (!string.IsNullOrEmpty(projectLibPath))
            {
                btnExecute.IsEnabled = true;
                btnSelectLibrary.IsEnabled = false;
                tbMessage.Text = "library is selected";
                return;
            }
            tbMessage.Text = "Path for library wrong. Pls select new path for library";

        }
        private void BtnExecute_Click(object sender, RoutedEventArgs e)
        {
            startTime = DateTime.Now;
            tbMessage.Text = "Starting of execution ";
            Common.CreateNewFolder(exportPath);
            Common.CreateNewFolder(sourcePath);
            Common.CreateNewFolder(templatePath);            

            bool isVible = chbVisible.IsChecked == false || chbVisible.IsChecked == null ? false : true;
            bool isClose = chbClose.IsChecked == false || chbVisible.IsChecked == null ? false : true;
            bool isSave = chbSave.IsChecked == false || chbSave.IsChecked == null ? false : true;
            bool isCompile = chbCompile.IsChecked == false || chbCompile.IsChecked == null ? false : true;
            if (_tiaprj == null)
            {
                _tiaprj = new TIAPrj(isVible);
            }

            foreach (var plc in _excelprj.BlocksStruct)
            {
                //InitTIAPrj(projectPath);

                plcSoftware = _tiaprj.GetPLC(plc.namePLC);
                AddValueToDataBlock(plcSoftware, plc);

                CteateUserConstants(plcSoftware, plc);
                tbMessage.Text = "Created user constants ";

                CreateEqConstants(plcSoftware, plc);
                tbMessage.Text = "Created equipments constants ";

                CteateTagsFromFile(plcSoftware, plc);
                tbMessage.Text = "Created tags for symbol tables ";

                ConnectLib(plcSoftware, plc, projectLibPath);
                tbMessage.Text = "Loaded/Updated project library from global library ";

                UpdateSupportBlocks(plcSoftware, plc, projectLibPath);
                tbMessage.Text = "Updated support blocks from global library ";

                UpdateTypeBlocks(plcSoftware, plc, projectLibPath);
                tbMessage.Text = "Updated types blocks from project/global library ";

                CreateInstanceBlocks(plcSoftware, plc);
                tbMessage.Text = "Created instances DBs ";

                CreateTemplateFCFromExcel(plcSoftware, plc);
                tbMessage.Text = "Created FCs for call instanced DBs ";

                EditFCFromExcelCallAllBlocks(plcSoftware, plc, closeProject: isClose, saveProject: isSave, compileProject: isCompile);
                tbMessage.Text = "Created FCs for templates ";
            }
            tbMessage.Text = "Execution is completed";
            btnConnectToOpened.IsEnabled = false;
            btnOpenTIA.IsEnabled = false;
            btnReadExcel.IsEnabled = true;
            btnExecute.IsEnabled = false;
            chbCreateTags.IsEnabled = true;
            chbCreateInsDB.IsEnabled = true;
            chbCreateFC.IsEnabled = true;
            if ((!isVible || isClose) && _tiaprj != null)
            {
                _tiaprj.DisposeTIA();
            }
            _tiaprj = null;
            _excelprj = null;

            DateTime finishTime = DateTime.Now;
            Trace.WriteLine(finishTime - startTime);

            Log.Close();
            
        }
  
        public  void CloseWindows()
        {
            try
            {
                _excelprj.CloseExcelFile();
                _excelprj = null;
                _tiaprj.DisposeTIA();
                _tiaprj=null;
            }
            catch (Exception)
            {

            }
        }
       
    }
}
