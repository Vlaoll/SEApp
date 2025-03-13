using Siemens.Engineering.Library;
using Siemens.Engineering.SW.Blocks;
using Siemens.Engineering.SW.Tags;
using seConfSW.Domain.Models;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;

namespace seConfSW.Services
{
    public class TiaService
    {
        private TIAPrj _tiaprj;
        private string _msg = string.Empty;
        private string _sourceDBPath = @"Samples\sourceDB\";
        private string _exportPath = @"Samples\export\";
        private string _sourcePath = @"samples\source\";
        private string _templatePath = @"samples\template\";
        private string _projectPath = null;
        private string _projectLibPath = null;
        private string _sourceTagPath = @"samples\tag\";

        public bool InitTIAPrj(string projectPath, out string message)
        {
            message = _msg;
            if (Directory.Exists(_exportPath)) Directory.Delete(_exportPath, true);
            if (Directory.Exists(_sourcePath)) Directory.Delete(_sourcePath, true);
            if (Directory.Exists(_templatePath)) Directory.Delete(_templatePath, true);

            Directory.CreateDirectory(_exportPath);
            Directory.CreateDirectory(_sourcePath);
            Directory.CreateDirectory(_templatePath);

            string filename;
            try
            {
                DateTime temp = DateTime.Now;
                message = temp + " :[General]Start to initialization TIA project";
                Trace.WriteLine(message);

                _tiaprj.StartTIA();
                filename = new FileInfo(projectPath).FullName;
                _tiaprj.OpenProject(filename);
                _tiaprj.ConnectTIA();

                temp = DateTime.Now;
                message = temp + " :[General]Finished to connect to project";
                _msg = message;
                return true;
            }
            catch (Exception ex)
            {
                message = "[General]" + ex.Message;
                Trace.WriteLine(message);
                _msg = message;
                return false;
            }
        }

        public bool ConnectToOpenedTiaProject(bool isVisible, out string message)
        {
            message = _msg;
            _tiaprj = new TIAPrj(isVisible);
            bool isNext = _tiaprj.ConnectTIA() == 2;
            message = isNext ? "Connecting is successful" : "Connecting to TIA project was wrong";
            _msg = message;
            return isNext;
        }

        public bool OpenTiaProject(bool isVisible, out string message)
        {
            message = _msg;
            _tiaprj = new TIAPrj(isVisible);
            if (_tiaprj.ConnectTIA() == 0)
            {
                if (string.IsNullOrEmpty(_projectPath))
                {
                    _projectPath = _tiaprj.SearchProject();
                }

                if (!string.IsNullOrEmpty(_projectPath))
                {
                    if (_tiaprj.ConnectTIA() == 2)
                    {
                        message = "TIA project is opened and connected";
                        _msg = message;
                        return true;
                    }
                }
                message = "TIA didn't selected or wrong path";
                _msg = message;
                return false;
            }
            message = "Please close TIA Portal";
            _msg = message;
            return false;
        }

        public bool SelectLibrary(out string message)
        {
            message = _msg;
            if (string.IsNullOrEmpty(_projectLibPath))
            {
                _projectLibPath = _tiaprj.SearchLibrary();
            }

            if (!string.IsNullOrEmpty(_projectLibPath))
            {
                message = "Library is selected";
                _msg = message;
                return true;
            }
            message = "Path for library wrong";
            _msg = message;
            return false;
        }

        public bool AddValueToDataBlock(Siemens.Engineering.SW.PlcSoftware plcSoftware, dataPLC dataPLC, bool closeProject = false, bool saveProject = false, bool compileProject = false)
        {
            List<dataDataBlockValue> dataDBValues = new List<dataDataBlockValue>();
            List<dataSupportBD> dataSupportBD = new List<dataSupportBD>();
            Dictionary<string, StringBuilder> DictDBBase = new Dictionary<string, StringBuilder>();
            Dictionary<string, Dictionary<string, string>> DictDBValues = new Dictionary<string, Dictionary<string, string>>();
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
                        if (!dataSupportBD.Any(item => item.name == equipment.name))
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
                    var filename = _tiaprj.GenerateSourceBlock(plcSoftware, DB.group + "." + DB.name, _sourcePath);
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

                        using (var sw = new StreamWriter(_sourcePath + DB.name + ".db"))
                        {
                            foreach (var item in codeSB.ToString().Replace('\r', ' ').Split('\n'))
                            {
                                sw.WriteLine(item);
                            }
                            filename = _sourcePath + DB.name + ".db";
                        }
                        _tiaprj.ClearSource(plcSoftware);
                        _tiaprj.ImportSource(plcSoftware, DB.name, filename);
                        _tiaprj.ChangeBlockNumber(plcSoftware, DB.name, DB.number, _tiaprj.GenerateBlock(plcSoftware, DB.group));
                        _msg = "[General]Created empty source for: " + DB.name;
                        Trace.WriteLine(_msg);
                    }
                    if (filename != string.Empty)
                    {
                        var tempFilename = System.IO.Path.GetTempFileName();
                        using (var sr = new StreamReader(filename))
                        {
                            using (var sw = new StreamWriter(tempFilename))
                            {
                                string line;
                                while ((line = sr.ReadLine()) != null)
                                {
                                    if (line.Contains("STRUCT") || line.Contains("VAR RETAIN"))
                                    {
                                        string[] str;
                                        line = sr.ReadLine();
                                        while (line != null && (!line.Contains("END_VAR") && !line.Contains("END_STRUCT;")))
                                        {
                                            str = Array.Empty<string>();
                                            str = line.Replace(" ", "").Split(':');
                                            if (str.Length > 1 && str.Length < 3)
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
                                    else if (line.Contains("BEGIN"))
                                    {
                                        break;
                                    }
                                    else
                                    {
                                        sw.WriteLine(line);
                                    }
                                }
                                if (DB.isRetain)
                                {
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
                        File.Delete(filename);
                        File.Move(tempFilename, filename);
                        File.Delete(tempFilename);

                        _tiaprj.ClearSource(plcSoftware);
                        _tiaprj.ImportSource(plcSoftware, DB.name, filename);
                        _tiaprj.ChangeBlockNumber(plcSoftware, DB.name, DB.number, _tiaprj.GenerateBlock(plcSoftware, DB.group));
                        _msg = "[General]Generated new source for: " + DB.name;
                        Trace.WriteLine(_msg);
                    }
                }
                return false; // Исходный код возвращает false в try-блоке
            }
            catch (Exception)
            {
                return false;
            }
        }

        public bool CreateInstanceBlocks(Siemens.Engineering.SW.PlcSoftware plcSoftware, dataPLC dataPLC, bool closeProject = false, bool saveProject = false, bool compileProject = false)
        {
            try
            {
                DateTime temp = DateTime.Now;
                _msg = temp + "[General]Start creating instance blocks";
                Trace.WriteLine(_msg);
                foreach (dataBlock blockDB in dataPLC.instanceDB)
                {
                    int numberDB = blockDB.number != 0 ? (blockDB.number - 1) * 20 + 1000 : 0;
                    if (!dataPLC.Equipment.Where(type => type.typeEq == blockDB.typeEq).FirstOrDefault().isExtended)
                    {
                        _tiaprj.CreateInstanceDB(plcSoftware, blockDB.name, numberDB, blockDB.instanceOfName, _sourceDBPath, blockDB.group, blockDB.typeEq);
                    }
                    else
                    {
                        foreach (dataExtSupportBlock item in dataPLC.Equipment.Where(type => type.typeEq == blockDB.typeEq).FirstOrDefault().dataExtSupportBlock)
                        {
                            if (item.variant.Count == 0 || blockDB.variant.Intersect(item.variant).Any())
                            {
                                string name = Common.ModifyString(item.name, blockDB.excelData);
                                _tiaprj.CreateInstanceDB(plcSoftware, name, numberDB + item.number, item.instanceOfName, _sourceDBPath, blockDB.group, blockDB.typeEq);
                            }
                        }
                    }
                }
                if (compileProject) { _tiaprj.Compile(dataPLC.namePLC); }
                if (saveProject) { _tiaprj.SaveProject(); }
                if (closeProject) { _tiaprj.CloseProject(); }
                return true;
            }
            catch (Exception ex)
            {
                _msg = "[General]" + ex.Message;
                Trace.WriteLine(_msg);
                return false;
            }
        }

        public bool EditFCFromExcelCallAllBlocks(Siemens.Engineering.SW.PlcSoftware plcSoftware, dataPLC dataPLC, ExcelDataReader excelprj, bool closeProject = false, bool saveProject = false, bool compileProject = false)
        {
            StringBuilder io;
            StringBuilder param;
            string[] code;
            var sources = new Dictionary<string, string>();

            DateTime temp = DateTime.Now;
            _msg = temp + "[General]Start editing FCs";
            Trace.WriteLine(_msg);

            try
            {
                sources.Clear();
                List<dataBlock> orderDataExcelListBlocks = dataPLC.instanceDB.Where(element => dataPLC.Equipment.Where(type => type.typeEq == element.typeEq).FirstOrDefault().isExtended == false)
                    .OrderBy(order => order.number)
                    .ToList();
                foreach (dataBlock data in orderDataExcelListBlocks)
                {
                    io = null;
                    io = new StringBuilder();
                    param = null;
                    param = new StringBuilder();
                    code = null;

                    if (data.typeEq.Length > 0 && data.nameEq.Length > 0)
                    {
                        dataEq EqType = excelprj.BlocksStruct
                            .Where(plc => plc.namePLC == dataPLC.namePLC)
                            .First()
                            .Equipment
                            .Where(eq => eq.typeEq == data.typeEq)
                            .First();

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
                            _tiaprj.CreateFC(plcSoftware, dateFC.name, dateFC.number, blockString, _sourcePath, groupName: dateFC.group);
                        }

                        sources.Add(data.nameFC, _tiaprj.GenerateSourceBlock(plcSoftware, dateFC.group + "." + dateFC.name, _sourcePath));
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
                if (compileProject) { _tiaprj.Compile(dataPLC.namePLC); }
                if (saveProject) { _tiaprj.SaveProject(); }
                if (closeProject) { _tiaprj.CloseProject(); }

                return true;
            }
            catch (Exception ex)
            {
                _msg = "[General]" + ex.Message;
                Trace.WriteLine(_msg);
                return false;
            }
        }

        public bool CreateTemplateFCFromExcel(Siemens.Engineering.SW.PlcSoftware plcSoftware, dataPLC dataPLC, bool closeProject = false, bool saveProject = false, bool compileProject = false)
        {
            var sources = new Dictionary<string, string>();
            var files = new Dictionary<string, string>();
            string[] code;
            if (Directory.Exists(_templatePath)) Directory.Delete(_templatePath, true);
            Directory.CreateDirectory(_templatePath);

            DateTime temp = DateTime.Now;
            _msg = temp + "[General]Start tempales FCs";
            Trace.WriteLine(_msg);

            try
            {
                List<dataBlock> orderDataExcelListBlocks = dataPLC.instanceDB.Where(element => dataPLC.Equipment.Where(type => type.typeEq == element.typeEq).FirstOrDefault().isExtended == true)
                    .OrderBy(order => order.number)
                    .ToList();

                var Eqs = dataPLC.Equipment.Where(typeEq => typeEq.isExtended == true);
                foreach (dataEq type in Eqs)
                {
                    dataLibrary template = type.FB.FirstOrDefault();
                    if (template != null) // Изменено с GetType() != null на проверку null
                    {
                        (PlcBlockGroup group, string nameBlock) = _tiaprj.GetBlockGroup(plcSoftware, template.group + "." + template.name);

                        PlcBlock plcBlock = group != null ? group.Blocks.Find(nameBlock) : null;
                        if (plcBlock != null)
                        {
                            _tiaprj.GenerateSourceBlock(plcSoftware, template.group + "." + template.name, _templatePath);
                        }
                    }
                }
                foreach (dataBlock data in orderDataExcelListBlocks)
                {
                    string filename = new FileInfo(_templatePath + data.instanceOfName + ".scl").FullName;
                    string eqFileName = new FileInfo(_templatePath + data.group + "." + "FC_Call-" + data.typeEq + ".scl").FullName;
                    code = null;
                    if (!files.ContainsKey(data.group + "." + "FC_Call-" + data.typeEq))
                    {
                        files.Add(data.group + "." + "FC_Call-" + data.typeEq, eqFileName);
                    }

                    using (var sr = new StreamReader(filename))
                    {
                        using (var sw = new StreamWriter(eqFileName, true))
                        {
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
                                    if ((lineNext = sr.ReadLine()).Contains("isVariant") && lineNext != null)
                                    {
                                        var n = lineNext.IndexOf("isVariant") + 12;
                                        List<string> str = lineNext.Remove(0, n).Split(',').ToList();
                                        var isVariantExist = data.variant.Intersect(str).ToList();
                                        if (!isVariantExist.Any())
                                        {
                                            while ((line = sr.ReadLine()) != "\tEND_REGION" && line != null) { }
                                        }
                                        else
                                        {
                                            sw.WriteLine(Common.ModifyString(line, data.excelData, "@"));
                                            sw.WriteLine(lineNext);
                                        }
                                    }
                                    else
                                    {
                                        sw.WriteLine(Common.ModifyString(line, data.excelData, "@"));
                                        sw.WriteLine(Common.ModifyString(lineNext, data.excelData, "@"));
                                    }
                                }
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
                            _tiaprj.CreateFC(plcSoftware, dateFC.name, dateFC.number, blockString, _sourcePath, groupName: dateFC.group);
                        }

                        sources.Add(data.nameFC, _tiaprj.GenerateSourceBlock(plcSoftware, dateFC.group + "." + dateFC.name, _sourcePath));
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

                if (compileProject) { _tiaprj.Compile(dataPLC.namePLC); }
                if (saveProject) { _tiaprj.SaveProject(); }
                if (closeProject) { _tiaprj.CloseProject(); }

                return true;
            }
            catch (Exception ex)
            {
                _msg = "[General]" + ex.Message;
                Trace.WriteLine(_msg);
                return false;
            }
        }

        public bool CteateTagTableFromFile(Siemens.Engineering.SW.PlcSoftware plcSoftware, dataPLC dataPLC, string groupName = "@Eq_TagTables", bool closeProject = false, bool saveProject = false, bool compileProject = false)
        {
            try
            {
                DateTime temp = DateTime.Now;
                _msg = temp + " :[General]Start creating tags";
                Trace.WriteLine(_msg);

                List<dataImportTag> sources = new List<dataImportTag>();
                List<dataExcistTag> existTags = new List<dataExcistTag>();

                if (Directory.Exists(_sourceTagPath)) Directory.Delete(_sourceTagPath, true);
                Directory.CreateDirectory(_sourceTagPath);
                string filename;
                foreach (dataBlock instance in dataPLC.instanceDB)
                {
                    var tags = dataPLC.Equipment.Where(item => item.typeEq == instance.typeEq).First().dataTag;
                    foreach (dataTag tag in tags)
                    {
                        if (tag.adress != "")
                        {
                            filename = new FileInfo(_sourceTagPath + tag.table + ".xml").FullName;
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

                    using (var sw = new StreamWriter(source.path, true))
                    {
                        sw.WriteLine(source.code.ToString());
                    }
                    _tiaprj.ImportTagTable(plcSoftware, source.path);
                }
                temp = DateTime.Now;
                _msg = temp + " :[General]Finished creating tags";
                Trace.WriteLine(_msg);
                if (compileProject) { _tiaprj.Compile(dataPLC.namePLC); }
                if (saveProject) { _tiaprj.SaveProject(); }
                if (closeProject) { _tiaprj.CloseProject(); }
                return true;
            }
            catch (Exception ex)
            {
                _msg = "[General]" + ex.Message;
                Trace.WriteLine(_msg);
                return false;
            }
        }

        public bool CteateTagsFromFile(Siemens.Engineering.SW.PlcSoftware plcSoftware, dataPLC dataPLC, string groupName = "@Eq_TagTables", bool closeProject = false, bool saveProject = false, bool compileProject = false)
        {
            try
            {
                DateTime temp = DateTime.Now;
                _msg = temp + " :[General]Start creating tags";
                Trace.WriteLine(_msg);

                List<dataImportTag> sources = new List<dataImportTag>();
                List<dataExcistTag> existTags = new List<dataExcistTag>();

                if (Directory.Exists(_sourceTagPath)) Directory.Delete(_sourceTagPath, true);
                Directory.CreateDirectory(_sourceTagPath);
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
                                filename = new FileInfo(_sourceTagPath + tag.table + ".xml").FullName;
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

                    using (var sw = new StreamWriter(source.path, true))
                    {
                        sw.WriteLine(source.code.ToString());
                    }
                    _tiaprj.ImportTags(plcSoftware, source.path, source.table);
                }
                temp = DateTime.Now;
                _msg = temp + " :[General]Finished creating tags";
                Trace.WriteLine(_msg);
                if (compileProject) { _tiaprj.Compile(dataPLC.namePLC); }
                if (saveProject) { _tiaprj.SaveProject(); }
                if (closeProject) { _tiaprj.CloseProject(); }
                return true;
            }
            catch (Exception ex)
            {
                _msg = "[General]" + ex.Message;
                Trace.WriteLine(_msg);
                return false;
            }
        }

        public bool CteateUserConstants(Siemens.Engineering.SW.PlcSoftware plcSoftware, dataPLC dataPLC, bool closeProject = false, bool saveProject = false, bool compileProject = false)
        {
            try
            {
                DateTime temp = DateTime.Now;
                _msg = temp + " :[General]Start to creating user constants";
                Trace.WriteLine(_msg);

                _tiaprj.CreateUserConstant(plcSoftware, dataPLC.userConstant);

                temp = DateTime.Now;
                _msg = temp + " :[General]Finished to creating user constants";
                Trace.WriteLine(_msg);

                if (compileProject) { _tiaprj.Compile(dataPLC.namePLC); }
                if (saveProject) { _tiaprj.SaveProject(); }
                if (closeProject) { _tiaprj.CloseProject(); }
                return true;
            }
            catch (Exception ex)
            {
                _msg = "[General]" + ex.Message;
                Trace.WriteLine(_msg);
                return false;
            }
        }

        public bool CreateTags(Siemens.Engineering.SW.PlcSoftware plcSoftware, dataPLC dataPLC, bool closeProject = false, bool saveProject = false, bool compileProject = false)
        {
            try
            {
                DateTime temp = DateTime.Now;
                _msg = temp + " [General]Start creating tags";
                Trace.WriteLine(_msg);
                _tiaprj.CreateTag(plcSoftware, dataPLC);

                temp = DateTime.Now;
                _msg = temp + " [General]Finished creating tags";
                Trace.WriteLine(_msg);

                if (compileProject) { _tiaprj.Compile(dataPLC.namePLC); }
                if (saveProject) { _tiaprj.SaveProject(); }
                if (closeProject) { _tiaprj.CloseProject(); }
                return true;
            }
            catch (Exception ex)
            {
                _msg = "[General]" + ex.Message;
                Trace.WriteLine(_msg);
                return false;
            }
        }

        public bool CreateEqConstants(Siemens.Engineering.SW.PlcSoftware plcSoftware, dataPLC dataPLC, bool closeProject = false, bool saveProject = false, bool compileProject = false)
        {
            try
            {
                DateTime temp = DateTime.Now;
                _msg = temp + " [General]Start creating equipment constants";
                Trace.WriteLine(_msg);

                foreach (dataBlock blockDB in dataPLC.instanceDB)
                {
                    List<userConstant> listOfConstants = dataPLC.Equipment.Where(type => type.typeEq == blockDB.typeEq).FirstOrDefault().dataConstant;
                    if (listOfConstants != null)
                        _tiaprj.CreateUserConstant(plcSoftware, listOfConstants, blockDB.excelData, blockDB.nameEq);
                }
                temp = DateTime.Now;
                _msg = temp + " [General]Finished creating equipment constants";
                Trace.WriteLine(_msg);

                if (compileProject) { _tiaprj.Compile(dataPLC.namePLC); }
                if (saveProject) { _tiaprj.SaveProject(); }
                if (closeProject) { _tiaprj.CloseProject(); }
                return true;
            }
            catch (Exception ex)
            {
                _msg = "[General]" + ex.Message;
                Trace.WriteLine(_msg);
                return false;
            }
        }

        public bool ConnectLib(Siemens.Engineering.SW.PlcSoftware plcSoftware, dataPLC dataPLC, string libraryPath, bool closeProject = false, bool saveProject = false, bool compileProject = false)
        {
            string fileNameLibrary;
            try
            {
                _msg = "[General]Start updating project library";
                Trace.WriteLine(_msg);
                fileNameLibrary = new FileInfo(libraryPath).FullName;
                _tiaprj.UpdatePrjLibraryFromGlobal(fileNameLibrary);
                if (compileProject) { _tiaprj.Compile(dataPLC.namePLC); }
                if (saveProject) { _tiaprj.SaveProject(); }
                if (closeProject) { _tiaprj.CloseProject(); }
                return true;
            }
            catch (Exception ex)
            {
                _msg = "[General]" + ex.Message;
                Trace.WriteLine(_msg);
                return false;
            }
        }

        public bool UpdateSupportBlocks(Siemens.Engineering.SW.PlcSoftware plcSoftware, dataPLC dataPLC, string libraryPath, bool closeProject = false, bool saveProject = false, bool compileProject = false)
        {
            string fileNameLibrary;
            try
            {
                _msg = "[General]Start updating support blocks from library";
                Trace.WriteLine(_msg);
                fileNameLibrary = new FileInfo(libraryPath).FullName;

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
                        else if (item.type == "UDT" && item.isType)
                        {
                            _tiaprj.GenerateUDTFromLibrary(plcSoftware, "Common", item.name, item.group);
                        }
                    }
                    catch (Exception ex)
                    {
                        _msg = "[General]" + ex.Message + " :" + item.name;
                        Trace.WriteLine(_msg);
                        return false;
                    }
                }
                globalLibrary.Close();

                if (compileProject) { _tiaprj.Compile(dataPLC.namePLC); }
                if (saveProject) { _tiaprj.SaveProject(); }
                if (closeProject) { _tiaprj.CloseProject(); }
                return true;
            }
            catch (Exception ex)
            {
                _msg = "[General]" + ex.Message;
                Trace.WriteLine(_msg);
                return false;
            }
        }

        public bool UpdateTypeBlocks(Siemens.Engineering.SW.PlcSoftware plcSoftware, dataPLC dataPLC, string libraryPath, bool closeProject = false, bool saveProject = false, bool compileProject = false, bool clearProjectLibrary = false)
        {
            string fileNameLibrary;
            try
            {
                _msg = "[General]Start updating types blocks from project library";
                Trace.WriteLine(_msg);
                fileNameLibrary = new FileInfo(libraryPath).FullName;
                UserGlobalLibrary globalLibrary = _tiaprj.OpenLibrary(fileNameLibrary);

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
                            _msg = "[General]" + ex.Message + " :" + block.name;
                            Trace.WriteLine(_msg);
                            return false;
                        }
                    }
                }
                globalLibrary.Close();

                if (clearProjectLibrary) { _tiaprj.CleanUpMainLibrary(); }
                if (compileProject) { _tiaprj.Compile(dataPLC.namePLC); }
                if (saveProject) { _tiaprj.SaveProject(); }
                if (closeProject) { _tiaprj.CloseProject(); }
                return true;
            }
            catch (Exception ex)
            {
                _msg = "[General]" + ex.Message;
                Trace.WriteLine(_msg);
                return false;
            }
        }

        private string DeleteBlockRegion(string filename, string eqName)
        {
            var tempFilename = System.IO.Path.GetTempFileName();
            try
            {
                using (var sr = new StreamReader(filename))
                {
                    using (var sw = new StreamWriter(tempFilename))
                    {
                        string line;
                        bool exit = false;
                        while ((line = sr.ReadLine()) != null)
                        {
                            int numLine = line.Trim().Length;
                            if (line.Contains("REGION Call") && line.Contains(eqName))
                            {
                                exit = true;
                                _msg = "[General]Delete Region for block: " + eqName;
                                Trace.WriteLine(_msg);
                            }
                            if (!line.Contains("REGION " + eqName) && !exit)
                            {
                                sw.WriteLine(line);
                            }
                            else
                            {
                                exit = line.Contains("END_REGION") ? false : true;
                            }
                        }
                    }
                }
                File.Delete(filename);
                File.Move(tempFilename, filename);
                File.Delete(tempFilename);
            }
            catch (Exception e)
            {
                _msg = "[General]Exception: " + e.Message;
                Trace.WriteLine(_msg); // Добавлено логирование, как в исходнике
            }
            return filename;
        }

        private bool AddBlockRegion(string filename, string[] newBlock, string eqName)
        {
            var tempFilename = System.IO.Path.GetTempFileName();
            using (var sr = new StreamReader(filename))
            {
                using (var sw = new StreamWriter(tempFilename))
                {
                    string line;
                    while ((line = sr.ReadLine()) != null)
                    {
                        if (!line.Contains("END_FUNCTION"))
                        {
                            sw.WriteLine(line);
                        }
                        else
                        {
                            foreach (var item in newBlock)
                            {
                                sw.WriteLine(item);
                            }
                            _msg = "[General]Add Region for block: " + eqName;
                            Trace.WriteLine(_msg);
                        }
                    }
                }
            }
            File.Delete(filename);
            File.Move(tempFilename, filename);
            File.Delete(tempFilename);
            return true;
        }

        public void DisposeTia()
        {
            try
            {
                _tiaprj?.DisposeTIA();
                _tiaprj = null;
            }
            catch (Exception)
            {
                // Исключение игнорируется, как в исходном коде
            }
        }

        public Siemens.Engineering.SW.PlcSoftware GetPLC(string plcName)
        {
            return _tiaprj.GetPLC(plcName);
        }

        public string GetMessage()
        {
            return _msg;
        }

        public string GetProjectLibPath()
        {
            return _projectLibPath;
        }
    }
}