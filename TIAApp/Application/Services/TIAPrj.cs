using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Reflection;
using System.Diagnostics;
using Microsoft.Win32;
using Siemens.Engineering;
using Siemens.Engineering.Compiler;
using Siemens.Engineering.Hmi;
using Siemens.Engineering.HW;
using Siemens.Engineering.HW.Features;
using Siemens.Engineering.SW;
using Siemens.Engineering.SW.Blocks;
using Siemens.Engineering.SW.Types;
using Siemens.Engineering.SW.ExternalSources;
using Siemens.Engineering.Library;
using Siemens.Engineering.Library.Types;
using Siemens.Engineering.SW.Tags;
using Siemens.Engineering.Library.MasterCopies;
using seConfSW.Domain.Models;

namespace seConfSW
{
    public class TIAPrj
    {
        private static TiaPortalProcess _tiaProcess;
        private string message = string.Empty;
        private bool checkUI;
        private List<dataBlock> dataPrjListFB = new List<dataBlock>();

        public TiaPortal MyTiaPortal
        {
            get; set;
        }
        public Project MyProject
        {
            get; set;
        }
        public string Message
        {
            get { return message; }

        }
        public List<dataBlock> DataPrjListFB
        {
            get { return dataPrjListFB; }

        }

        public TIAPrj(bool checkUI = false)
        {
            AppDomain CurrentDomain = AppDomain.CurrentDomain;
            CurrentDomain.AssemblyResolve += new ResolveEventHandler(MyResolver);
            this.checkUI = checkUI;
        }

        private static Assembly MyResolver(object sender, ResolveEventArgs args)
        {
            int index = args.Name.IndexOf(',');
            if (index == -1)
            {
                return null;
            }
            string name = args.Name.Substring(0, index);

            RegistryKey filePathReg = Registry.LocalMachine.OpenSubKey(
                "SOFTWARE\\Siemens\\Automation\\Openness\\18.0\\PublicAPI\\18.0.0.0");

            if (filePathReg == null)
                return null;

            object oRegKeyValue = filePathReg.GetValue(name);
            if (oRegKeyValue != null)
            {
                string filePath = oRegKeyValue.ToString();

                string path = filePath;
                string fullPath = Path.GetFullPath(path);
                if (File.Exists(fullPath))
                {
                    return Assembly.LoadFrom(fullPath);
                }
            }

            return null;
        }

        public void StartTIA()
        {
            if (checkUI)
            {
                MyTiaPortal = new TiaPortal(TiaPortalMode.WithUserInterface);
                _tiaProcess = TiaPortal.GetProcesses()[0];
                message = "[TIA]" +  "TIA Portal started  with user interface";
                Trace.WriteLine(message);
            }
            else
            {
                MyTiaPortal = new TiaPortal(TiaPortalMode.WithoutUserInterface);
                message = "[TIA]" +  "TIA Portal started without user interface";
                Trace.WriteLine(message);
            }
        }
        public void DisposeTIA() 
        {
            if (MyTiaPortal != null)
            {
                MyTiaPortal.Dispose();
                message = "[TIA]" +  "TIA Portal disposed";
                Trace.WriteLine(message);

                _tiaProcess.Dispose();
                MyTiaPortal = null;
            }            
        }
        public PlcSoftware GetPLC (string plcName)
        {
            try
            {
                DeviceItem deviceItem = null;
                foreach (var devices in MyProject.Devices)
                {
                    if (devices.DeviceItems.Any(item => item.Name.Contains(plcName)))
                    {
                        deviceItem = devices.DeviceItems.Where(item => item.Name.Contains(plcName)).First();
                    }
                }               
                   
                if (deviceItem != null)
                {
                    SoftwareContainer softwareContainer = ((IEngineeringServiceProvider)deviceItem).GetService<SoftwareContainer>();
                    if (softwareContainer != null)
                    {
                        message = "[TIA]" +  "Get PLC from project: " + plcName;
                        Trace.WriteLine(message);
                        return softwareContainer.Software as PlcSoftware;  
                    }
                }
            }
            catch (Exception ex)
            {
                message = "[TIA]" +  "Error: " + ex.Message;
                Trace.WriteLine(message);
                return null;
            }
            return null;
        }



        public string SearchProject(string filter = "TIA Portal v19 |*.ap19")
        {
            OpenFileDialog fileSearch = new OpenFileDialog();
            

            fileSearch.Multiselect = false;
            fileSearch.ValidateNames = true;
            fileSearch.DereferenceLinks = false; // Will return .lnk in shortcuts.
            fileSearch.Filter = filter;
            fileSearch.RestoreDirectory = true;
            fileSearch.InitialDirectory = Environment.CurrentDirectory;
            fileSearch.Filter = filter;
            
            fileSearch.ShowDialog();

            string ProjectPath = fileSearch.FileName.ToString();

            if (string.IsNullOrEmpty(ProjectPath) == false)
            {
                StartTIA();
                if (OpenProject(ProjectPath))
                {
                    message = "[TIA]" + "Open project " + ProjectPath;
                    Trace.WriteLine(message);                    
                } 
            }
            return ProjectPath;
        }
        public string SearchLibrary(string filter = "TIA Library v19 |*.al19")
        {
            OpenFileDialog fileSearch = new OpenFileDialog();

            fileSearch.Multiselect = false;
            fileSearch.ValidateNames = true;
            fileSearch.DereferenceLinks = false; // Will return .lnk in shortcuts.
            fileSearch.Filter = filter;
            fileSearch.RestoreDirectory = true;
            fileSearch.InitialDirectory = Environment.CurrentDirectory;
            fileSearch.Filter = filter;

            fileSearch.ShowDialog();
            string ProjectPath = fileSearch.FileName.ToString();
            return ProjectPath;            
        }
        public bool OpenProject(string ProjectPath)
        {
            try
            {
                message = "[TIA]" +  "Opening project " + ProjectPath;
                Trace.WriteLine(message);
                MyProject = MyTiaPortal.Projects.Open(new FileInfo(ProjectPath));
                message = "[TIA]" +  "Project " + ProjectPath + " opened";
                Trace.WriteLine(message);
                return true;
            }
            catch (Exception ex)
            {
                message = "[TIA]" +  "Error while opening project" + ex.Message;
                Trace.WriteLine(message);
                return false;
            }
        }
        public int ConnectTIA()
        {
            message = "[TIA]" +  "Trying to connect to project";
            Trace.WriteLine(message);

            IList<TiaPortalProcess> processes = TiaPortal.GetProcesses();
            switch (processes.Count)
            {
                case 0:
                    message = "[TIA]" +  "No running instance of TIA Portal was found!";
                    Trace.WriteLine(message);
                    return 0;
                case 1:
                    _tiaProcess = processes[0];
                    MyTiaPortal = _tiaProcess.Attach();
                    if (MyTiaPortal.GetCurrentProcess().Mode == TiaPortalMode.WithUserInterface)
                    { checkUI = true; }
                    else
                    {checkUI = false;}
                    if (MyTiaPortal.Projects.Count <= 0)
                    {
                        message = "[TIA]" +  "No TIA Portal Project was found!";
                        Trace.WriteLine(message);
                        return 1;
                    }
                    MyProject = MyTiaPortal.Projects[0];
                    message = "[TIA]" +  "Connected to Project: " + MyProject.Name;
                    Trace.WriteLine(message);
                    return 2;                                    
                default:
                    message = "[TIA]" +  "More than one running instance of TIA Portal was found!";
                    Trace.WriteLine(message);
                    return 3;
            }            
        }
        private void DisConnectTIA()
        {
            
        }
        public void SaveProject()
        {
            message = "[TIA]" +  "Saving the project ";
            Trace.WriteLine(message);
            MyProject.Save();
            message = "[TIA]" +  "Project saved";
            Trace.WriteLine(message);
        }
        public void CloseProject()
        {
            message = "[TIA]" +  "Closeing the project";
            Trace.WriteLine(message);

            MyProject.Close();
            message = "[TIA]" +  "Project closed";
            Trace.WriteLine(message);
        }
        public UserGlobalLibrary OpenLibrary(string LibraryPath)
        {
            try
            {
                UserGlobalLibrary globalLibrary = MyTiaPortal.GlobalLibraries.Open(new FileInfo(LibraryPath), OpenMode.ReadWrite);
                message = "[TIA]" + "Open global library: " + LibraryPath;
                Trace.WriteLine(message);
                return globalLibrary;
            }
            catch (Exception ex)
            {
                message = "[TIA]" + "Error while opening global library" + ex.Message;
                Trace.WriteLine(message);
                return null;
            }

        }
        public void UpdatePrjLibraryFromGlobal(string LibraryPath, bool CloseLibrary = true)
        //Open global library
        {
            try
            {
                message = "[TIA]" +  "Updating project library ";
                Trace.WriteLine(message);
                UserGlobalLibrary globalLibrary = OpenLibrary(LibraryPath);

                ProjectLibrary projectLibrary = MyProject.ProjectLibrary;
                var systemTypeFolder = new[] { globalLibrary.TypeFolder };                

                globalLibrary.UpdateLibrary(systemTypeFolder, projectLibrary);
                message = "[TIA]" +  "Project Library is updated";
                Trace.WriteLine(message);

                if (CloseLibrary)
                {
                    globalLibrary.Close();
                    message = "[TIA]" + "Global library is closed";
                    Trace.WriteLine(message);
                }
            }
            catch (Exception ex)
            {
                message = "[TIA]" +  "Error while opening library" + ex.Message;
                Trace.WriteLine(message);
            }
        }
        public bool CopyBlocksFromMasterCopyFolder(PlcSoftware plcSoftware, UserGlobalLibrary globalLibrary, string blockName,  string groupName = null)
        //Copy blocks from master copies
        {
            string[] hierarchyBlock = blockName.Split('.');
            int numHierarchyBlock = hierarchyBlock.Length;                       

            try
            {
                PlcBlockSystemGroup systemGroup = plcSoftware.BlockGroup;
                string plcName = plcSoftware.Name;

                if (hierarchyBlock.Length > 1)
                {
                    MasterCopyUserFolder userFolder = null;
                    MasterCopy userBlock = null;
                    userFolder = globalLibrary.MasterCopyFolder.Folders.Where(item => item.Name == hierarchyBlock[0]).FirstOrDefault();
                    hierarchyBlock = hierarchyBlock.Where((val, idx) => idx != 0).ToArray();

                    for (int i = 1; i < numHierarchyBlock; i++)
                    {
                        try
                        {
                            if (hierarchyBlock.Length > 1)
                            {
                                userFolder = userFolder.Folders.Where(item => item.Name == hierarchyBlock[0]).FirstOrDefault();
                                hierarchyBlock = hierarchyBlock.Where(val => val != hierarchyBlock[0]).ToArray();
                            }
                            else
                            {
                                userBlock = userFolder.MasterCopies.Where(item => item.Name.Equals(hierarchyBlock[0], StringComparison.OrdinalIgnoreCase)).FirstOrDefault();
                            }
                        }
                        catch (Exception ex)
                        {
                            message = "[TIA]" + "Error while seaching DB in MC" + ex.Message;
                            Trace.WriteLine(message);
                        }
                    }
                    if (userBlock != null)
                    {
                        if (groupName != null)
                        {
                            PlcBlockUserGroup myCreatedGroup = CreateBlockGroup(plcSoftware, groupName);
                            if (myCreatedGroup.Blocks.Find(userBlock.Name) == null)
                            {
                                myCreatedGroup.Blocks.CreateFrom(userBlock);
                                message = "[TIA]" + "Created block for Master Copy Folder: " + userBlock.Name;
                                Trace.WriteLine(message);
                            }
                            return true;
                        }
                        else
                        {
                            if (systemGroup.Blocks.Find(userBlock.Name) == null)
                            {
                                systemGroup.Blocks.CreateFrom(userBlock);
                                message = "[TIA]" + "Created block for Master Copy Folder: " + userBlock.Name;
                                Trace.WriteLine(message);
                            }
                            return true;
                        }
                    }
                    return false;
                }
                else { return false; }
            }
            catch (Exception ex)
            {
                message = "[TIA]" +  "Error: " + ex.Message;
                Trace.WriteLine(message);
                return false;
            }
        }
        public bool CopyUDTFromMasterCopyFolder(PlcSoftware plcSoftware, UserGlobalLibrary globalLibrary, string blockName, string groupName = null)
        //Copy blocks from master copies
        {
            string[] hierarchy = blockName.Split('.');
            int numHierarchy = hierarchy.Length;            

            try
            { 
                PlcTypeSystemGroup systemGroup = plcSoftware.TypeGroup;                       

                if (hierarchy.Length > 1)
                {
                    MasterCopyUserFolder userFolder = null;
                    MasterCopy userBlock = null;
                    userFolder = globalLibrary.MasterCopyFolder.Folders.Where(item => item.Name == hierarchy[0]).FirstOrDefault();
                    hierarchy = hierarchy.Where(val => val != hierarchy[0]).ToArray();

                    for (int i = 1; i < numHierarchy; i++)
                    {

                        try
                        {
                            if (hierarchy.Length > 1)
                            {
                                userFolder = userFolder.Folders.Where(item => item.Name == hierarchy[0]).FirstOrDefault();
                                hierarchy = hierarchy.Where(val => val != hierarchy[0]).ToArray();
                            }
                            else
                            {
                                userBlock = userFolder.MasterCopies.Where(item => item.Name == hierarchy[0]).FirstOrDefault();
                            }
                        }
                        catch (Exception ex)
                        {
                            message = "[TIA]" + "Error while seaching UDT in MC" + ex.Message;
                            Trace.WriteLine(message);
                        }
                    }
                    if (userBlock != null)
                    {
                        if (userBlock != null)
                        {
                            if (groupName != null)
                            {
                                PlcTypeUserGroup myCreatedGroup = systemGroup.Groups.Find(groupName);
                                if (myCreatedGroup == null)
                                {
                                    myCreatedGroup = systemGroup.Groups.Create(groupName);
                                    message = "[TIA]" + "Create group for UDT: " + groupName;
                                    Trace.WriteLine(message);
                                }
                                if (myCreatedGroup.Types.Find(userBlock.Name) == null)
                                {
                                    myCreatedGroup.Types.CreateFrom(userBlock);
                                    message = "[TIA]" + "Create block for Master Copy Folder: " + userBlock.Name;
                                    Trace.WriteLine(message);
                                }
                                return true;
                            }
                            else
                            {
                                if (systemGroup.Types.Find(userBlock.Name) == null)
                                {
                                    systemGroup.Types.CreateFrom(userBlock);
                                    message = "[TIA]" + "Create block for Master Copy Folder: " + userBlock.Name;
                                    Trace.WriteLine(message);
                                }
                                return true;
                            }
                        }
                        return false;
                    }
                    else { return false; }
                }
                else { return false; }

            }
            catch (Exception ex)
            {
                message = "[TIA]" +  "Error: " + ex.Message;
                Trace.WriteLine(message);
                return false;
            }
        }


        public void Compile(string devname)
        {
            bool found = false;

            foreach (Device device in MyProject.Devices)
            {
                DeviceItemComposition deviceItemAggregation = device.DeviceItems;
                foreach (DeviceItem deviceItem in deviceItemAggregation)
                {
                    if (deviceItem.Name == devname || device.Name == devname)
                    {
                        SoftwareContainer softwareContainer = deviceItem.GetService<SoftwareContainer>();
                        if (softwareContainer != null)
                        {
                            if (softwareContainer.Software is PlcSoftware)
                            {
                                PlcSoftware controllerTarget = softwareContainer.Software as PlcSoftware;
                                if (controllerTarget != null)
                                {
                                    found = true;
                                    ICompilable compiler = controllerTarget.GetService<ICompilable>();
                                    message = "[TIA]" +  "Compiling of " + controllerTarget.Name ;
                                    Trace.WriteLine(message);

                                    CompilerResult result = compiler.Compile();
                                    message = "[TIA]" +  "Compiling of " + controllerTarget.Name + ": State: " + result.State + " / Warning Count: " + result.WarningCount + " / Error Count: " + result.ErrorCount;
                                    Trace.WriteLine(message);
                                }
                            }
                            if (softwareContainer.Software is HmiTarget)
                            {
                                HmiTarget hmitarget = softwareContainer.Software as HmiTarget;
                                if (hmitarget != null)
                                {
                                    found = true;
                                    ICompilable compiler = hmitarget.GetService<ICompilable>();
                                    message = "[TIA]" +  "Compiling of " + hmitarget.Name;
                                    Trace.WriteLine(message);

                                    CompilerResult result = compiler.Compile();
                                    message = "[TIA]" +  "Compiling of " + hmitarget.Name + ": State: " + result.State + " / Warning Count: " + result.WarningCount + " / Error Count: " + result.ErrorCount;
                                    Trace.WriteLine(message);
                                }

                            }
                        }
                    }
                }
            }
            if (found == false)
            {
                message = "[TIA]" +  "Didn't find device with name " + devname;
                Trace.WriteLine(message);
            }
        }
        public void AddHW(string nameDevice, string orderNo, string version)
        {
            string MLFB = "OrderNumber:" + orderNo + "/" + version;
            string name = nameDevice;
            string devname = "station" + nameDevice;
            bool found = false;
            foreach (Device device in MyProject.Devices)
            {
                DeviceItemComposition deviceItemAggregation = device.DeviceItems;
                foreach (DeviceItem deviceItem in deviceItemAggregation)
                {
                    if (deviceItem.Name == devname || device.Name == devname)
                    {
                        SoftwareContainer softwareContainer = deviceItem.GetService<SoftwareContainer>();
                        if (softwareContainer != null)
                        {
                            if (softwareContainer.Software is PlcSoftware)
                            {
                                PlcSoftware controllerTarget = softwareContainer.Software as PlcSoftware;
                                if (controllerTarget != null)
                                {
                                    found = true;
                                }
                            }
                            if (softwareContainer.Software is HmiTarget)
                            {
                                HmiTarget hmitarget = softwareContainer.Software as HmiTarget;
                                if (hmitarget != null)
                                {
                                    found = true;
                                }

                            }
                        }
                    }
                }
            }
            if (found == true)
            {
                message = "[TIA]" +  "Device " + nameDevice + " already exists";
                Trace.WriteLine(message);
            }
            else
            {
                Device deviceName = MyProject.Devices.CreateWithItem(MLFB, name, devname);
                message = "[TIA]" +  "Add Device Name: " + name + " with Order Number: " + orderNo + " and Firmware Version: " + version;
                Trace.WriteLine(message);
            }
        }

        public string GenerateSourceBlock(PlcSoftware plcSoftware, string blockName, string path, GenerateOptions generateOption = GenerateOptions.None)
        {
            // exports all blocks and with all their dependencies(e.g. called blocks, used DBs or UDTs)
            // as ASCII text into the provided source file
            string filename = string.Empty;
            string plcName = plcSoftware.Name;
            try
            {                               
                var blocks = new List<PlcBlock>();
                (PlcBlockGroup group, string name) = GetBlockGroup(plcSoftware, blockName);
                PlcBlock plcBlock = group.Blocks.Find(name);                        
                if (plcBlock != null)
                {
                    blocks.Add(plcBlock);
                    string blockType = string.Empty;

                    switch (plcBlock.ProgrammingLanguage)
                    {
                        case ProgrammingLanguage.Undef:
                            break;
                        case ProgrammingLanguage.STL:
                            blockType = ".awl";
                            break;
                        case ProgrammingLanguage.SCL:
                            blockType = ".scl";
                            break;
                        case ProgrammingLanguage.DB:
                            blockType = ".db";
                            break;

                        default:
                            break;
                    }
                    if (blockType != string.Empty)
                    {
                        filename = new FileInfo(path + plcBlock.Name + blockType).FullName;
                        if (File.Exists(filename))
                        {
                            File.Delete(filename);
                            message = "[TIA]" +  "Delete file: " + filename;
                            Trace.WriteLine(message);
                        }
                        var fileInfo = new FileInfo(filename);
                        PlcExternalSourceSystemGroup systemGroup = plcSoftware.ExternalSourceGroup;
                        if (blocks.Count > 0)
                        {
                            systemGroup.GenerateSource(blocks, fileInfo, generateOption);
                            message = "[TIA]" +  "Generate source for block: " + blockName;
                            Trace.WriteLine(message);
                        }
                    }
                }                         
            }
            catch (Exception ex)
            {
                message = "[TIA]" +  "Error: " + ex.Message;
                Trace.WriteLine(message);
                return null;
            }
            return filename;
        }
        public string GenerateSourceUDT(PlcSoftware plcSoftware, string typeName, string path)
        {
            // exports all blocks and with all their dependencies(e.g. called blocks, used DBs or UDTs)
            // as ASCII text into the provided source file
            string filename = string.Empty;
            string plcName = plcSoftware.Name;
            try
            {
                var blocks = new List<PlcType>();
                PlcType plcBlock = plcSoftware.TypeGroup.Types.Find(typeName);
                blocks.Add(plcBlock);                                

                if (plcBlock != null)
                {
                    filename = new FileInfo(path + plcBlock.Name + ".udt").FullName;
                    if (File.Exists(filename))
                    {
                        File.Delete(filename);
                        message = "[TIA]" +  "Delete file: " + filename;
                        Trace.WriteLine(message);
                    }
                    //var fileInfo = new FileInfo(filename);
                    PlcExternalSourceSystemGroup systemGroup = plcSoftware.ExternalSourceGroup;
                    if (blocks.Count > 0)
                    {
                        systemGroup.GenerateSource(blocks, new FileInfo(filename), GenerateOptions.WithDependencies);
                        message = "[TIA]" +  "Generate source from block: " + plcName;
                        Trace.WriteLine(message);
                    }
                }
            }
            catch (Exception ex)
            {
                message = "[TIA]" +  "Error: " + ex.Message;
                Trace.WriteLine(message);
                return null;
            }
            return filename;
        }
        public PlcBlockGroup GenerateBlock(PlcSoftware plcSoftware,  string groupName = null  )
        // Creates a block from an external source file
        {
            string[] hierarchy = groupName.Split('.');
            int numHierarchy = hierarchy.Length;
            string plcName = plcSoftware.Name;
            try
            {
                foreach (PlcExternalSource plcExternalSource in plcSoftware.ExternalSourceGroup.ExternalSources)
                {
                    PlcBlockUserGroup groupUser;
                    if (numHierarchy > 0)
                    {
                        groupUser = CreateBlockGroup(plcSoftware, hierarchy[0]);                       

                        hierarchy = hierarchy.Where(val => val != hierarchy[0]).ToArray();
                        for (int i = 1; i < numHierarchy; i++)
                        {
                            groupUser = CreateBlockGroup(plcSoftware, hierarchy[0], groupUser.Groups);                            
                        }
                                                                
                        plcExternalSource.GenerateBlocksFromSource(groupUser, GenerateBlockOption.KeepOnError);
                        message = "[TIA]" +  "Generated blocks from source: " + plcExternalSource.Name;
                        Trace.WriteLine(message);
                        return groupUser;
                    }
                    else
                    {
                        plcExternalSource.GenerateBlocksFromSource(GenerateBlockOption.KeepOnError);
                        message = "[TIA]" +  "Generated blocks from source: " + plcExternalSource.Name;
                        Trace.WriteLine(message);
                        return plcSoftware.BlockGroup;
                    }
                                                               
                }
            }
            catch (Exception ex)
            {
                message = "[TIA]" +  "Error: " + ex.Message;
                Trace.WriteLine(message);
                return null;
            }
            return null;
        }
        public void ClearSource(PlcSoftware plcSoftware)
        {
            // Clear  external source files
            string plcName = plcSoftware.Name;
            try
            {                        
                while (plcSoftware.ExternalSourceGroup.ExternalSources.Count > 0)
                {
                    message = "[TIA]" +  "Clear source: " + plcSoftware.ExternalSourceGroup.ExternalSources[0].Name;
                    Trace.WriteLine(message);

                    plcSoftware.ExternalSourceGroup.ExternalSources[0].Delete();
                }                        
               
            }
            catch (Exception ex)
            {
                message = "[TIA]" +  "Error: " + ex.Message;
                Trace.WriteLine(message);                
            }
        }
        public void ImportSource(PlcSoftware plcSoftware, string blockName, string path, string type)
        // Creates a block from a AWL, SCL, DB or UDT file
        {
            string plcName = plcSoftware.Name;
            try
            {   
                PlcExternalSource externalSource = plcSoftware.ExternalSourceGroup.ExternalSources.Find(blockName);
                if (externalSource != null)
                {
                    externalSource.Delete();
                    message = "[TIA]" +  "Delete all sources ";
                    Trace.WriteLine(message);
                }

                string filename = new FileInfo(path + blockName + type).FullName;
                externalSource = plcSoftware.ExternalSourceGroup.ExternalSources.CreateFromFile(blockName, filename);
                message = "[TIA]" +  "Import source: " + blockName;
                Trace.WriteLine(message);
  
            }
            catch (Exception ex)
            {
                message = "[TIA]" +  "Error: " + ex.Message;
                Trace.WriteLine(message);
            }

        }
        public void ImportSource(PlcSoftware plcSoftware, string blockName, string path)
        // Creates a block from a AWL, SCL, DB or UDT file
        {
           try
            { 
                PlcExternalSource externalSource = plcSoftware.ExternalSourceGroup.ExternalSources.Find(blockName);
                if (externalSource != null)
                {
                    externalSource.Delete();
                    message = "[TIA]" +  "Delete all sources ";
                    Trace.WriteLine(message);
                }

                string filename = new FileInfo(path).FullName;
                externalSource = plcSoftware.ExternalSourceGroup.ExternalSources.CreateFromFile(blockName, filename);
                message = "[TIA]" +  "Import source: " + blockName;
                Trace.WriteLine(message);

            }
            catch (Exception ex)
            {
                message = "[TIA]" +  "Error: " + ex.Message;
                Trace.WriteLine(message);
            }
        }
        public void GenerateBlockFromLibrary(PlcSoftware plcSoftware, string pathName, string typeName, string groupName = null)
        {          
            ProjectLibrary projectLibrary = MyProject.ProjectLibrary;
            var group = GetTypeGroup(projectLibrary.TypeFolder, pathName);
            object typeCode ;
            CodeBlockLibraryTypeVersion typeBlock;
            PlcTypeLibraryTypeVersion typePlcType;

            try
            {
                if (group != null)
            { typeCode = group.Types.Find(typeName).Versions.First((version) => version.IsDefault) ; }
            else { return; }
                switch (typeCode.GetType().Name)
                {
                    case "CodeBlockLibraryTypeVersion":
                        typeBlock = typeCode as CodeBlockLibraryTypeVersion;
                        PlcBlockGroup blockGroup = plcSoftware.BlockGroup;
                        if (groupName != null)
                        {
                            PlcBlockUserGroup groupUser = CreateBlockGroup(plcSoftware, groupName);
                            message = "[TIA]" + "Create group folder " + groupName;
                            Trace.WriteLine(message);
                            groupUser.Blocks.CreateFrom(typeBlock);
                            message = "[TIA]" + "Import block " + typeName;
                            Trace.WriteLine(message);
                        }
                        else
                        {
                            blockGroup.Blocks.CreateFrom(typeBlock);
                            message = "[TIA]" + "Import block " + typeName;
                            Trace.WriteLine(message);
                        }
                        break;
                    case "PlcTypeLibraryTypeVersion":
                        typePlcType = typeCode as PlcTypeLibraryTypeVersion;
                        PlcTypeSystemGroup typeGroup = plcSoftware.TypeGroup;
                        if (groupName != null)
                        {
                            PlcTypeUserGroup groupUser = CreateTypeGroup(plcSoftware, groupName);
                            message = "[TIA]" + "Create tape's group folder " + groupName;
                            Trace.WriteLine(message);
                            groupUser.Types.CreateFrom(typePlcType);
                            message = "[TIA]" + "Import type " + typeName;
                            Trace.WriteLine(message);
                        }
                        else
                        {
                            typeGroup.Types.CreateFrom(typePlcType);
                            message = "[TIA]" + "Import type " + typeName;
                            Trace.WriteLine(message);
                        }
                        break;
                    default:
                        break;
                }
            string plcName = plcSoftware.Name; 
            }
            catch (Exception ex)
            {
                message = "[TIA]" +  "Error: " + ex.Message;
                Trace.WriteLine(message);
            }
        }
        public void GenerateUDTFromLibrary(PlcSoftware plcSoftware, string typeName, string blockName, string groupName = null)
        {
            ProjectLibrary projectLibrary = MyProject.ProjectLibrary;
            PlcTypeLibraryTypeVersion typeCodeBlock = projectLibrary.TypeFolder.Folders.Find(typeName).Folders.Find("UDTs").Types.Find(blockName).Versions.First((version) => version.IsDefault) as PlcTypeLibraryTypeVersion;
            string plcName = plcSoftware.Name;

            try
            {
                PlcTypeSystemGroup blockGroup = plcSoftware.TypeGroup;
                if (groupName != null)
                {
                    PlcTypeUserGroup myCreatedGroup = blockGroup.Groups.Find(groupName);
                    if (myCreatedGroup == null)
                    {
                        myCreatedGroup = blockGroup.Groups.Create(groupName);
                        message = "[TIA]" + "Create group for UDT: " + groupName;
                        Trace.WriteLine(message);
                    }
                    if (myCreatedGroup.Types.Find(blockName) == null)
                    {
                        myCreatedGroup.Types.CreateFrom(typeCodeBlock);
                        message = "[TIA]" + "Import UDT " + blockName;
                        Trace.WriteLine(message);
                    }                    
                }                
                else
                {
                    if (blockGroup.Types.Find(blockName) == null)
                    {
                        blockGroup.Types.CreateFrom(typeCodeBlock);
                        message = "[TIA]" + "Import UDT " + blockName;
                        Trace.WriteLine(message);
                    }                    
                }
            }
            catch (Exception ex)
            {
                message = "[TIA]" + "Error: " + ex.Message;
                Trace.WriteLine(message);
            }
        }
        public void CleanUpMainLibrary()
        // Clean up main library
        {
            message = "[TIA]" +  "Clean up main library";
            Trace.WriteLine(message);
            ProjectLibrary projectLibrary = MyProject.ProjectLibrary;        
            projectLibrary.CleanUpLibrary(projectLibrary.TypeFolder.Folders, CleanUpMode.DeleteUnusedTypes);
        }
        
        public void ExportBlock(PlcSoftware plcSoftware, string blockName, string path)
        {        
            try
            {                     
                PlcBlock plcBlock = plcSoftware.BlockGroup.Blocks.Find(blockName);
                string filename = new FileInfo(path + plcBlock.Name + ".xml").FullName;
                                
                if (File.Exists(filename))
                {
                    File.Delete(filename);
                    message = "[TIA]" +  "Delete file" + filename;
                    Trace.WriteLine(message);
                }
                               
                plcBlock.Export(new FileInfo(filename), ExportOptions.WithDefaults);
                message = "[TIA]" +  "Export block" + blockName;
                Trace.WriteLine(message);

            }
            catch (Exception ex)
            {
                message = "[TIA]" +  "Error: " + ex.Message;
                Trace.WriteLine(message);
            }
        }
        public void ImportBlock(PlcSoftware plcSoftware, string blockName, string path, string groupName = null)
        {
            string plcName = plcSoftware.Name;
            try
            {
                PlcBlockGroup blockGroup = plcSoftware.BlockGroup;

                string filename = new FileInfo(path + blockName + ".xml").FullName;
                IList<PlcBlock> blocks;
                if (groupName != null)
                {
                    PlcBlockUserGroup groupUser = CreateBlockGroup(plcSoftware, groupName);
                    blocks = groupUser.Blocks.Import(new FileInfo(filename), ImportOptions.Override, SWImportOptions.IgnoreMissingReferencedObjects);
                    message = "[TIA]" +  "Import block " + blockName;
                    Trace.WriteLine(message);
                }
                else
                {
                    blocks = blockGroup.Blocks.Import(new FileInfo(filename), ImportOptions.Override, SWImportOptions.IgnoreMissingReferencedObjects);
                    message = "[TIA]" +  "Import block " + blockName;
                    Trace.WriteLine(message);
                }                      
                
            }
            catch (Exception ex)
            {
                message = "[TIA]" +  "Error: " + ex.Message;
                Trace.WriteLine(message);
            }
        }
        public void DeleteBlock(PlcSoftware plcSoftware, string blockName)
        // Delete a block
        {
            string plcName = plcSoftware.Name;
            try
            {
                foreach (PlcBlockUserGroup group in plcSoftware.BlockGroup.Groups)
                {
                    if (group.Blocks.Find(blockName) != null)
                    {
                        group.Blocks.Find(blockName).Delete();
                        message = "[TIA]" +  "Delete block: " + blockName;
                        Trace.WriteLine(message);
                    }
                }

            }
            catch (Exception ex)
            {
                message = "[TIA]" +  "Error: " + ex.Message;
                Trace.WriteLine(message);
            }
        }

        public void ChangeBlockNumber(PlcSoftware plcSoftware, string blockName, int number, PlcBlockGroup group)
        {
            string plcName = plcSoftware.Name;
            try
            {      
                if (group.Blocks.Find(blockName) != null)
                {
                    group.Blocks.Find(blockName).AutoNumber = false;
                    group.Blocks.Find(blockName).Number = number;
                    message = "[TIA]" +  "Change number of block: " + blockName + " to " + number;
                    Trace.WriteLine(message);
                }                                        
              
            }
            catch (Exception ex)
            {
                message = "[TIA]" +  "Error: " + ex.Message;
                Trace.WriteLine(message);
            }
        }
        public void CreateDB(PlcSoftware plcSoftware, string blockName, int number, string instanceOfName, string path, string groupName = null)
        {
            StringBuilder sb = new StringBuilder("DATA_BLOCK " + '"' + blockName + '"');
            sb.AppendLine("");
            sb.AppendLine('"' + instanceOfName + '"');
            sb.AppendLine("BEGIN");
            sb.AppendLine("END_DATA_BLOCK");
            string blockString = sb.ToString();
            
            try
            {
                string filename = new FileInfo(path + instanceOfName + "_instanceDB" + ".db").FullName;
                if (File.Exists(filename))
                {
                    File.Delete(filename);
                    message = "[TIA]" +  "Deleted file: " + filename;
                    Trace.WriteLine(message);
                }
                using (StreamWriter sw = File.CreateText(filename))
                {
                    sw.WriteLine(blockString);                   
                    message = "[TIA]" +  "Created Instance DB: " + blockName;
                    Trace.WriteLine(message);
                }
                ClearSource(plcSoftware);
                ImportSource(plcSoftware, instanceOfName + "_instanceDB", path, ".db");
                var group = GenerateBlock(plcSoftware, groupName);
                if (number !=0)
                {
                    ChangeBlockNumber(plcSoftware, blockName, number, group);
                }
                
                message = "[TIA]" +  "Created Instance DB: " + blockName;
                Trace.WriteLine(message);

                ClearSource(plcSoftware);
            }
            catch (Exception ex)
            {
                message = "[TIA]" +  "Error: " + ex.Message;
                Trace.WriteLine(message);
            }
        }
        public void CreateFC(PlcSoftware plcSoftware, string blockName, int number, string blockString, string path, string codeType = ".scl", string groupName = null )
        {            
            try
            {
                string filename = new FileInfo(path + blockName + codeType).FullName;
                if (File.Exists(filename))
                {
                    File.Delete(filename);
                    message = "[TIA]" +  "Deleted file: " + blockName;
                    Trace.WriteLine(message);
                }
                using (StreamWriter sw = File.CreateText(filename))
                {
                    sw.WriteLine(blockString); 
                    message = "[TIA]" +  "Created source file: " + blockName;
                    Trace.WriteLine(message);
                }
                ClearSource(plcSoftware);
                ImportSource(plcSoftware, blockName, path, codeType);
                DeleteBlock(plcSoftware, blockName);
                var group = GenerateBlock(plcSoftware, groupName);
                if (number != 0)
                {
                    ChangeBlockNumber(plcSoftware, blockName, number, group);
                }
                
                message = "[TIA]" +  "Created FC: " + blockName;
                Trace.WriteLine(message);

                ClearSource(plcSoftware);
            }
            catch (Exception ex)
            {
                message = "[TIA]" +  "Error: " + ex.Message;
                Trace.WriteLine(message);
            }
        }
        public void CreateFC(PlcSoftware plcSoftware, string blockName, int number, string path, string groupName = null)
        {
            try
            {
                ClearSource(plcSoftware);
                ImportSource(plcSoftware, blockName, path);
                DeleteBlock(plcSoftware, blockName);
                var group = GenerateBlock(plcSoftware, groupName);
                if (number != 0)
                {
                    ChangeBlockNumber(plcSoftware, blockName, number, group);
                }
                
                message = "[TIA]" +  "Created FC: " + blockName;
                Trace.WriteLine(message);

                ClearSource(plcSoftware);
            }
            catch (Exception ex)
            {
                message = "[TIA]" +  "Error: " + ex.Message;
                Trace.WriteLine(message);
            }
        }
        public void CreateInstanceDB(PlcSoftware plcSoftware, string blockName, int number, string instanceOfName, string path, string groupName, string groupEqName)
        {            
            try
            {
                DeleteBlock(plcSoftware, blockName);
                if (groupName != null)
                {
                    PlcBlockUserGroup groupUser = CreateBlockGroup(plcSoftware, groupName, groupEqName);
                    var instant = groupUser.Blocks.Find(blockName);
                    if (instant == null)
                    {                       
                        bool iaAutoNumber = number == 0 ? true : false;
                        int num = number == 0 ? 1 : number;
                        groupUser.Blocks.CreateInstanceDB(blockName, iaAutoNumber, number, instanceOfName);
                        message = "[TIA]" + "Created Instance DB: " + blockName;
                        Trace.WriteLine(message);
                    }                   
                }
                else
                {
                    var instant = plcSoftware.BlockGroup.Blocks.Find(blockName);
                    if (instant == null)
                    {
                        plcSoftware.BlockGroup.Blocks.CreateInstanceDB(blockName, false, number, instanceOfName);
                        message = "[TIA]" + "Created Instance DB: " + blockName;
                        Trace.WriteLine(message);
                    }
                   
                }
            }
            catch (Exception ex)
            {
                message = "[TIA]" +  "Error: " + ex.Message;
                Trace.WriteLine(message);
            }
        }
        public void CreateListFB(PlcSoftware plcSoftware, List<dataBlock> dataBlock)
        // Create list FB
        {
            string plcName = plcSoftware.Name;
            try
            {
                foreach (PlcBlockUserGroup group in plcSoftware.BlockGroup.Groups)
                {
                    if (group.Name.Contains("DB") && group.Name.Contains("@"))
                    {
                        foreach (var eqGroup in group.Groups)
                        {
                            foreach (var block in eqGroup.Blocks)
                            {
                                if (block is InstanceDB)
                                {
                                    InstanceDB bl = (InstanceDB)block;
                                    int pos = bl.Name.IndexOf("iDB");
                                    string[] eq = null;
                                    if (pos != -1)
                                    {
                                        eq = bl.Name.Substring(4).Split('|');
                                    }

                                    dataPrjListFB.Add(new dataBlock()
                                    {
                                        name = bl.Name,
                                        instanceOfName = bl.InstanceOfName,
                                        number = bl.Number,
                                        nameFC = dataBlock.Where(item => item.name == bl.Name.Substring(4)).First().nameFC,
                                        group = group.Name,
                                        typeEq = (pos != -1) && eq.Length > 1 ? eq[0] : "",
                                        nameEq = (pos != -1) && eq.Length > 1 ? eq[1] : "",
                                    });
                                    message = "[TIA]" +  "Added  " + bl.Name + " to the list of FB";
                                    Trace.WriteLine(message);
                                }
                            }
                        }
                    }
                }
                message = "[TIA]" +  "Created list for FB ";
                Trace.WriteLine(message);
            }
            catch (Exception ex)
            {
                message = "[TIA]" +  "Error: " + ex.Message;
                Trace.WriteLine(message);
            }
        }
        public PlcTypeUserGroup CreateTypeGroup(PlcSoftware plcSoftware, string groupName)
        //Creates a block group
        {
            string[] hierarchy = groupName.Split('.');
            int numHierarchy = hierarchy.Length;
            try
            {
                PlcTypeSystemGroup systemGroup = plcSoftware.TypeGroup;
                PlcTypeUserGroupComposition groupComposition = systemGroup.Groups;
                PlcTypeUserGroup myCreatedGroup = null;
                if (hierarchy.Length >= 1)
                {

                    for (int i = 1; i <= numHierarchy; i++)
                    {
                        if (groupComposition.Find(hierarchy[0]) == null)
                        {
                            myCreatedGroup = groupComposition.Create(hierarchy[0]);
                            message = "[TIA]" + "Create group for types: " + groupName;
                            Trace.WriteLine(message);
                            groupComposition = myCreatedGroup.Groups;
                            hierarchy = hierarchy.Where(val => val != hierarchy[0]).ToArray();
                        }
                        else
                        {
                            myCreatedGroup = groupComposition.Find(hierarchy[0]);
                            groupComposition = myCreatedGroup.Groups;
                            hierarchy = hierarchy.Where(val => val != hierarchy[0]).ToArray();
                        }
                    }
                    return myCreatedGroup;
                }
                return null;
            }
            catch (Exception ex)
            {
                message = "[TIA]" + "Error: " + ex.Message;
                Trace.WriteLine(message);
                return null;
            }
        }
        public PlcBlockUserGroup CreateBlockGroup(PlcSoftware plcSoftware, string groupName)
        //Creates a block group
        {
            string[] hierarchy = groupName.Split('.');
            int numHierarchy = hierarchy.Length;
            try
            {
                PlcBlockSystemGroup systemGroup = plcSoftware.BlockGroup;
                PlcBlockUserGroupComposition groupComposition = systemGroup.Groups;
                PlcBlockUserGroup myCreatedGroup = null;
                if (hierarchy.Length >= 1)
                {
                    
                    for (int i = 1; i <= numHierarchy; i++)
                    {
                        if (groupComposition.Find(hierarchy[0]) == null)
                        {
                            myCreatedGroup = groupComposition.Create(hierarchy[0]);
                            message = "[TIA]" + "Create group for blocks: " + groupName;
                            Trace.WriteLine(message);
                            groupComposition = myCreatedGroup.Groups;
                            hierarchy = hierarchy.Where(val => val != hierarchy[0]).ToArray();
                        }
                        else
                        {
                            myCreatedGroup = groupComposition.Find(hierarchy[0]);
                            groupComposition = myCreatedGroup.Groups;
                            hierarchy = hierarchy.Where(val => val != hierarchy[0]).ToArray();
                        }  
                    }
                    return myCreatedGroup;
                }
                return null;
            }
            catch (Exception ex)
            {
                message = "[TIA]" +  "Error: " + ex.Message;
                Trace.WriteLine(message);
                return null;
            }
        }
        public PlcBlockUserGroup CreateBlockGroup(PlcSoftware plcSoftware, string groupName, string groupEqName)
        //Creates a block group
        {
            try
            {
                PlcBlockSystemGroup systemGroup = plcSoftware.BlockGroup;
                PlcBlockUserGroupComposition groupComposition = systemGroup.Groups;

                if (groupComposition.Find(groupName) == null)
                {
                    PlcBlockUserGroup myCreatedGroup = groupComposition.Create(groupName);
                    message = "[TIA]" +  "Created group for blocks: " + groupName;
                    Trace.WriteLine(message);
                }
                PlcBlockUserGroupComposition groupEqComposition = groupComposition.Find(groupName).Groups;
                if (groupEqComposition.Find(groupEqName) == null)
                {
                    PlcBlockUserGroup myCreatedEqGroup = groupEqComposition.Create(groupEqName);
                    message = "[TIA]" +  "Created group for blocks: " + groupEqName;
                    Trace.WriteLine(message);
                }
                return groupEqComposition.Find(groupEqName);
            }

            catch (Exception ex)
            {
                message = "[TIA]" +  "Error: " + ex.Message;
                Trace.WriteLine(message);
                return null;
            }
        }
        public PlcBlockUserGroup CreateBlockGroup(PlcSoftware plcSoftware, string groupName, PlcBlockUserGroupComposition groupParent)
        //Creates a block group
        {
            try
            {             

                if (groupParent.Find(groupName) == null)
                {
                    PlcBlockUserGroup myCreatedGroup = groupParent.Create(groupName);
                    message = "[TIA]" +  "Create group for blocks: " + groupName;
                    Trace.WriteLine(message);
                }
                return groupParent.Find(groupName);
            }
            catch (Exception ex)
            {
                message = "[TIA]" +  "Error: " + ex.Message;
                Trace.WriteLine(message);
                return null;
            }
        }

        public PlcTagTable CreateTagTable(PlcSoftware plcSoftware, string groupName)
        //Creates a block group
        {
            string[] hierarchy = groupName.Split('.');
            int numHierarchy = hierarchy.Length;
            try
            {
                PlcTagTableSystemGroup systemGroup = plcSoftware.TagTableGroup;
                PlcTagTableUserGroupComposition groupComposition = systemGroup.Groups;
                PlcTagTableUserGroup myCreatedGroup = null;
                if (hierarchy.Length >= 1)
                {

                    for (int i = 1; i < numHierarchy; i++)
                    {
                        if (groupComposition.Find(hierarchy[0]) == null)
                        {
                            myCreatedGroup = groupComposition.Create(hierarchy[0]);
                            message = "[TIA]" + "Create group for Tag Tables: " + groupName;
                            Trace.WriteLine(message);
                            groupComposition = myCreatedGroup.Groups;
                            hierarchy = hierarchy.Where(val => val != hierarchy[0]).ToArray();
                        }
                        else
                        {
                            myCreatedGroup = groupComposition.Find(hierarchy[0]);
                            groupComposition = myCreatedGroup.Groups;
                            hierarchy = hierarchy.Where(val => val != hierarchy[0]).ToArray();
                        }
                    }
                    var table = myCreatedGroup.TagTables.Find(hierarchy[0]);
                    if (table == null)
                    {
                        return myCreatedGroup.TagTables.Create(hierarchy[0]);
                    }
                    else
                    {
                        return table;
                    }
                    
                }
                return null;
            }
            catch (Exception ex)
            {
                message = "[TIA]" + "Error: " + ex.Message;
                Trace.WriteLine(message);
                return null;
            }
        }
        public void  CreateTag(PlcSoftware plcSoftware,dataPLC dataPLC, string groupName = "@Eq_TagTables")
        //Creates a Tag Table group
        {
            try
            {
                PlcTagTableSystemGroup systemGroup = plcSoftware.TagTableGroup;
                PlcTagTableUserGroupComposition groupComposition = systemGroup.Groups;
                PlcTagTableUserGroup group = groupComposition.Find(groupName);
                if (group == null)
                {
                    group = groupComposition.Create(groupName);
                    message = "[TIA]" +  "Created group for symbol tables: " + groupName;
                    Trace.WriteLine(message);
                }

                foreach (dataBlock instance in dataPLC.instanceDB)
                {
                    foreach (dataTag tag in dataPLC.Equipment.Where(item => item.typeEq == instance.typeEq).First().dataTag)
                    {
                        if (tag.variant.Count == 0 || instance.variant.Intersect(tag.variant).Any())
                        {
                            if (tag.adress != "")
                            {
                                PlcTagTable TagTable = group.TagTables.Find(tag.table);
                                if (TagTable == null)
                                {
                                    TagTable = group.TagTables.Create(tag.table);
                                    message = "[TIA]" + "Created  symbol table: " + tag.table;
                                    Trace.WriteLine(message);
                                }
                                var link = Common.ModifyString(tag.link, instance.excelData);
                                if (TagTable.Tags.Find(link) == null)
                                {
                                    TagTable.Tags.Create(link, tag.type, tag.adress);
                                    message = "[TIA]" + "Created  tag: " + link;
                                    Trace.WriteLine(message);
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                message = "[TIA]" +  "Error: " + ex.Message;
                Trace.WriteLine(message);
            }
        }
        public void ImportTags(PlcSoftware plcSoftware, string filename,string table, string groupName = "@Eq_TagTables")
        //Import file with tags
        {
            try
            {
                FileInfo path = new FileInfo( new FileInfo(filename).FullName);
                PlcTagTableSystemGroup systemGroup = plcSoftware.TagTableGroup;
                
                PlcTagTableUserGroupComposition groupComposition = systemGroup.Groups;
                PlcTagTableUserGroup group = groupComposition.Find(groupName);
                if (group == null)
                {
                    group = groupComposition.Create(groupName);
                    message = "[TIA]" + "Created group for symbol tables: " + groupName;
                    Trace.WriteLine(message);
                }
                PlcTagTable TagTable = group.TagTables.Find(table);
                if (TagTable == null)
                {
                    TagTable = group.TagTables.Create(table);
                    message = "[TIA]" + "Created  symbol table: " + table;
                    Trace.WriteLine(message);
                }
                TagTable.Tags.Import(new FileInfo(filename), ImportOptions.None);                

                message = "[TIA]" + "Updated tags for table " + table;
                Trace.WriteLine(message);                
            }
            catch (Exception ex)
            {
                message = "[TIA]" + "Error: " + ex.Message;
                Trace.WriteLine(message);
               
            } 
        }        
        public void ImportTagTable(PlcSoftware plcSoftware, string filename, string groupName = "@Eq_TagTables")
        //Import file with tag table
        {
            try
            {
                FileInfo path = new FileInfo(new FileInfo(filename).FullName);
                PlcTagTableSystemGroup systemGroup = plcSoftware.TagTableGroup;

                PlcTagTableUserGroupComposition groupComposition = systemGroup.Groups;
                PlcTagTableUserGroup group = groupComposition.Find(groupName);
                if (group == null)
                {
                    group = groupComposition.Create(groupName);
                    message = "[TIA]" + "Created group for symbol tables: " + groupName;
                    Trace.WriteLine(message);
                }
                PlcTagTableComposition tagTableComposition = group.TagTables;                
                tagTableComposition.Import(path, ImportOptions.Override);
                
                message = "[TIA]" + "Created tags from file " + filename;
                Trace.WriteLine(message);

            }
            catch (Exception ex)
            {
                message = "[TIA]" + "Error: " + ex.Message;
                Trace.WriteLine(message);
                
            }
        }
        public (dataTag tag, bool isExist) FindTag(PlcSoftware plcSoftware, string tag, string table, string groupName = "@Eq_TagTables")
        //Find tag
        {
            dataTag tagData = new dataTag();
            try
            {
                PlcTagTableSystemGroup systemGroup = plcSoftware.TagTableGroup;
                
                PlcTagTableUserGroupComposition groupComposition = systemGroup.Groups;
                PlcTagTableUserGroup group = groupComposition.Find(groupName);
                if (group == null)
                {
                    return (tagData, false);
                }
                PlcTagTable TagTable = group.TagTables.Find(table);
                if (TagTable == null)
                {
                    return (tagData, false);
                }
                var t = TagTable.Tags.Find(tag);
                if (t == null)
                {
                    return (tagData, false);
                }
                tagData.link = t.Name;
                tagData.adress = t.LogicalAddress;
                tagData.type = t.DataTypeName;
                tagData.comment = t.Comment.ToString();
                return (tagData, true);
            }
            catch (Exception ex)
            {
                message = "[TIA]" + "Error: " + ex.Message;
                Trace.WriteLine(message);
                return (tagData, false);
            }
        }
        public void CreateUserConstant(PlcSoftware plcSoftware, List<userConstant> userConstant)
        //Creates a Tag Table group
        {
            try
            {
                PlcTagTableSystemGroup systemGroup = plcSoftware.TagTableGroup;
                var tagTableComposition = systemGroup.TagTables;
                PlcTagTable TagTable = tagTableComposition.Where(item => item.IsDefault == true).First();

                if (TagTable != null)
                {
                    if (userConstant.Count>0)
                    {
                        foreach (userConstant item in userConstant)
                        {
                            if (TagTable.UserConstants.Find(item.name) == null)
                            {
                                TagTable.UserConstants.Create(item.name, item.type, item.value);
                                message = "[TIA]" +  "Created user constant: " + item.name;
                                Trace.WriteLine(message);
                            }
                        }
                    }                     
                }
            }
            catch (Exception ex)
            {
                message = "[TIA]" +  "Error: " + ex.Message;
                Trace.WriteLine(message);
            }
        }
        public void CreateUserConstant(PlcSoftware plcSoftware, List<userConstant> userConstant, List<excelData> excelData,  string EqName,string groupName = "@Eq_TagTables")
        //Creates a Tag Table group
        {
            try
            {
                PlcTagTableSystemGroup systemGroup = plcSoftware.TagTableGroup;
                PlcTagTable TagTable = null;
                if (userConstant.Count > 0)
                {
                    List<userConstant> EqConstants = new List<userConstant>();
                    foreach (userConstant item in userConstant)
                    {
                        EqConstants.Add(new userConstant()
                        {
                            name = Common.ModifyString(item.name, excelData),
                            type = item.type,
                            value = Common.ModifyString(item.value, excelData),
                            table = item.table,
                        }
                            );
                    }
                    foreach (userConstant item in EqConstants)
                    {
                        if (string.IsNullOrWhiteSpace(item.table))
                        {
                            var tagTableComposition = systemGroup.TagTables;
                            TagTable = tagTableComposition.Where(gr => gr.IsDefault == true).First();
                        }
                        else
                        {
                            TagTable = CreateTagTable(plcSoftware, groupName + "." + item.table); 
                        }
                        if (TagTable.UserConstants.Find(item.name) == null)
                        {
                            TagTable.UserConstants.Create(item.name, item.type, item.value);

                        }
                    }
                    message = "[TIA]" + "Created equipment constants for: " + EqName;
                    Trace.WriteLine(message);
                }
               
            }
            catch (Exception ex)
            {
                message = "[TIA]" + "Error: " + ex.Message;
                Trace.WriteLine(message);
            }
        }

        public (PlcBlockGroup group, string name) GetBlockGroup(PlcSoftware plcSoftware, string blockName)
        // Get group and block name
        {
            string[] hierarchy = blockName.Split('.');
            int numHierarchy = hierarchy.Length;
            try
            {
                PlcBlockGroup group = plcSoftware.BlockGroup;

                if (hierarchy.Length > 1)
                {
                    for (int i = 1; i < numHierarchy; i++)
                    {
                        group = group.Groups.Where(item => item.Name == hierarchy[0]).FirstOrDefault();
                        hierarchy = hierarchy.Where(val => val != hierarchy[0]).ToArray();
                        if (group == null)
                        {
                            message = "[TIA]" +  "Wrong path for: " + hierarchy[hierarchy.Length - 1] +  " or missing";
                            Trace.WriteLine(message);
                            return (null, hierarchy[hierarchy.Length - 1]);
                        }
                    }
                    return (group, hierarchy[0]);
                }
                return (null, hierarchy[0]);
            }
            catch (Exception ex)
            {
                message = "[TIA]" +  "Error: " + ex.Message;
                Trace.WriteLine(message);
                return (null, null);
            }
        }
        public LibraryTypeUserFolder GetTypeGroup(LibraryTypeSystemFolder TypeFolder, string TypeName)
        // Get group and block name
        {
            //CodeBlockLibraryTypeVersion typeCodeBlock = projectLibrary.TypeFolder.Folders.Find(typeName).Folders.Find("Blocks").Types.Find(blockName).Versions.First((version) => version.IsDefault) as CodeBlockLibraryTypeVersion;

            string[] hierarchy = TypeName.Split('.');
            int numHierarchy = hierarchy.Length;
            try
            {
                LibraryTypeUserFolder group = null;

                if (hierarchy.Length > 1)
                {
                    
                    for (int i = 1; i <= numHierarchy; i++)
                    {
                        group = i==1?TypeFolder.Folders.Find(hierarchy[0]): group.Folders.Find(hierarchy[0]);
                        hierarchy = hierarchy.Where(val => val != hierarchy[0]).ToArray();
                        if (group == null)
                        {
                            message = "[TIA]" + "Wrong path for: " + hierarchy[hierarchy.Length - 1] + " or missing";
                            Trace.WriteLine(message);
                            return null;
                        }
                    }
                    return group;
                }
                return null;
            }
            catch (Exception ex)
            {
                message = "[TIA]" + "Error: " + ex.Message;
                Trace.WriteLine(message);
                return null;
            }
        }
    }    
}


