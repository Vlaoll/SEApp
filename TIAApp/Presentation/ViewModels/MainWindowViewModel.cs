using Microsoft.Extensions.Configuration;
using Microsoft.Win32;
using seConfSW.Services;
using System;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Runtime.CompilerServices;
using System.Windows.Input;

namespace seConfSW.Presentation.ViewModels
{
    public class MainWindowViewModel : INotifyPropertyChanged
    {
        private readonly ExcelService _excelService;
        private readonly TiaService _tiaService;
        private readonly IConfiguration _configuration;
        private string _message;
        private bool _canCheckPermission = true;
        private bool _canReadExcel = false;
        private bool _canConnectToOpened = false;
        private bool _canOpenTia = false;
        private bool _canSelectLibrary = false;
        private bool _canExecute = false;
        private bool _createTags = true;
        private bool _createInsDB = true;
        private bool _createFC = true;
        private bool _visibleTia = true;
        private bool _closeProject = false;
        private bool _compileProject = true;
        private bool _saveProject = true;
        private DateTime _startTime;
        private FileStream _log;

        public event PropertyChangedEventHandler PropertyChanged;

        public string Title { get; } = "SE TIA portal - constructor (V0.4.0)"; 

        public string Message
        {
            get => _message;
            set => SetProperty(ref _message, value);
        }

        public bool CanCheckPermission
        {
            get => _canCheckPermission;
            set => SetProperty(ref _canCheckPermission, value);
        }

        public bool CanReadExcel
        {
            get => _canReadExcel;
            set => SetProperty(ref _canReadExcel, value);
        }

        public bool CanConnectToOpened
        {
            get => _canConnectToOpened;
            set => SetProperty(ref _canConnectToOpened, value);
        }

        public bool CanOpenTia
        {
            get => _canOpenTia;
            set => SetProperty(ref _canOpenTia, value);
        }

        public bool CanSelectLibrary
        {
            get => _canSelectLibrary;
            set => SetProperty(ref _canSelectLibrary, value);
        }

        public bool CanExecute
        {
            get => _canExecute;
            set => SetProperty(ref _canExecute, value);
        }

        public bool CreateTags
        {
            get => _createTags;
            set => SetProperty(ref _createTags, value);
        }

        public bool CreateInsDB
        {
            get => _createInsDB;
            set => SetProperty(ref _createInsDB, value);
        }

        public bool CreateFC
        {
            get => _createFC;
            set => SetProperty(ref _createFC, value);
        }

        public bool VisibleTia
        {
            get => _visibleTia;
            set => SetProperty(ref _visibleTia, value);
        }

        public bool CloseProject
        {
            get => _closeProject;
            set => SetProperty(ref _closeProject, value);
        }

        public bool CompileProject
        {
            get => _compileProject;
            set => SetProperty(ref _compileProject, value);
        }

        public bool SaveProject
        {
            get => _saveProject;
            set => SetProperty(ref _saveProject, value);
        }

        public ICommand CheckPermissionCommand { get; }
        public ICommand ReadExcelCommand { get; }
        public ICommand ConnectToOpenedCommand { get; }
        public ICommand OpenTiaCommand { get; }
        public ICommand SelectLibraryCommand { get; }
        public ICommand ExecuteCommand { get; }

        public MainWindowViewModel(ExcelService excelService, TiaService tiaService, IConfiguration configuration)
        {
            _excelService = excelService;
            _tiaService = tiaService;
            _configuration = configuration;

            if (DateTime.Now > new DateTime(2025, 06, 30, 23, 59, 59))
            {
                CanCheckPermission = false;
            }

            CheckPermissionCommand = new RelayCommand(ExecuteCheckPermission, () => CanCheckPermission);
            ReadExcelCommand = new RelayCommand(ExecuteReadExcel, () => CanReadExcel);
            ConnectToOpenedCommand = new RelayCommand(ExecuteConnectToOpened, () => CanConnectToOpened);
            OpenTiaCommand = new RelayCommand(ExecuteOpenTia, () => CanOpenTia);
            SelectLibraryCommand = new RelayCommand(ExecuteSelectLibrary, () => CanSelectLibrary);
            ExecuteCommand = new RelayCommand(Execute, () => CanExecute);
        }

        private void ExecuteCheckPermission()
        {  
            System.Security.Principal.WindowsPrincipal principal = new System.Security.Principal.WindowsPrincipal(System.Security.Principal.WindowsIdentity.GetCurrent());
            if (!principal.IsInRole("Siemens TIA Openness"))
            {
                Message = "Add user to group: Siemens TIA Openness";
                return;
            }

            RegistryKey filePathReg = Registry.LocalMachine.OpenSubKey("SOFTWARE\\Siemens\\Automation\\Openness\\19.0\\PublicAPI\\19.0.0.0");
            if (filePathReg == null)
            {
                Message = "Missing or Incorrect version of TIA Portal. Must be TIA Portal v19";
                return;
            }
            Message = "All permissions are ok";
            CanReadExcel = true;
            CanCheckPermission = false;
        }

        private void StartTrace()
        {
            string logPath = @"log\";
            if (!Directory.Exists(logPath)) Directory.CreateDirectory(logPath);
            string LogPath = new FileInfo(logPath).FullName;
            string lodTime = DateTime.Now.ToString("yyyyMMdd'_'HHmm");
            _log = new FileStream(LogPath + "Log_" + lodTime + ".txt", FileMode.OpenOrCreate);
            Trace.Listeners.Add(new TextWriterTraceListener(_log));
            Trace.AutoFlush = true;
            Trace.WriteLine("Start log file: " + lodTime);
        }

        private void ExecuteReadExcel()
        {
            CanCheckPermission = false;
            _startTime = DateTime.Now;
            StartTrace();
            Message = "Start to read Excel file";
            if (_excelService.CreateExcelDB(out string msg))
            {
                Message = "Reading Excel file is completed. Select needed action for TIA project";
                CanConnectToOpened = true;
                CanOpenTia = true;
                CanReadExcel = false;
            }
            else
            {
                Message = "Reading Excel file is wrong.";
            }
        }

        private void ExecuteConnectToOpened()
        {
            VisibleTia = true;
            Message = "Try connect to opened TIA project";
            if (_tiaService.ConnectToOpenedTiaProject(VisibleTia, out string msg))
            {
                Message = "Connecting is successful. Pls select path for library";
                CanReadExcel = false;
                CanOpenTia = false;
                CanConnectToOpened = false;
                CanSelectLibrary = true;
            }
            else
            {
                Message = "Connecting to TIA project was wrong. Try again or open project";
            }
        }

        private void ExecuteOpenTia()
        {
            Message = "Try to open new TIA project";
            if (_tiaService.OpenTiaProject(VisibleTia, out string msg))
            {
                Message = "TIA project is opened and connected. Pls select path for library";
                CanReadExcel = false;
                CanConnectToOpened = false;
                CanOpenTia = false;
                CanSelectLibrary = true;
            }
            else
            {
                Message = msg;
            }
        }

        private void ExecuteSelectLibrary()
        {
            if (_tiaService.SelectLibrary(out string msg))
            {
                Message = "Library is selected";
                CanSelectLibrary = false;
                CanExecute = true;
            }
            else
            {
                Message = "Path for library wrong. Pls select new path for library";
            }
        }

        private void Execute()
        {
            _startTime = DateTime.Now;
            Message = "Starting of execution ";
            Common.CreateNewFolder(@"Samples\export\");
            Common.CreateNewFolder(@"samples\source\");
            Common.CreateNewFolder(@"samples\template\");

            var excelData = _excelService.GetExcelDataReader();
            foreach (var plc in excelData.BlocksStruct)
            {
                var plcSoftware = _tiaService.GetPLC(plc.namePLC);
                _tiaService.AddValueToDataBlock(plcSoftware, plc);

                _tiaService.CteateUserConstants(plcSoftware, plc);
                Message = "Created user constants ";

                _tiaService.CreateEqConstants(plcSoftware, plc);
                Message = "Created equipments constants ";

                _tiaService.CteateTagsFromFile(plcSoftware, plc);
                Message = "Created tags for symbol tables ";

                _tiaService.ConnectLib(plcSoftware, plc, _tiaService.GetProjectLibPath());
                Message = "Loaded/Updated project library from global library ";

                _tiaService.UpdateSupportBlocks(plcSoftware, plc, _tiaService.GetProjectLibPath());
                Message = "Updated support blocks from global library ";

                _tiaService.UpdateTypeBlocks(plcSoftware, plc, _tiaService.GetProjectLibPath());
                Message = "Updated types blocks from project/global library ";

                _tiaService.CreateInstanceBlocks(plcSoftware, plc);
                Message = "Created instances DBs ";

                _tiaService.CreateTemplateFCFromExcel(plcSoftware, plc);
                Message = "Created FCs for call instanced DBs ";

                _tiaService.EditFCFromExcelCallAllBlocks(plcSoftware, plc, excelData, CloseProject, SaveProject, CompileProject);
                Message = "Created FCs for templates ";
            }

            Message = "Execution is completed";
            CanConnectToOpened = false;
            CanOpenTia = false;
            CanReadExcel = true;
            CanExecute = false;

            if ((!VisibleTia || CloseProject))
            {
                _tiaService.DisposeTia();
                _excelService.CloseExcelFile();
            }

            DateTime finishTime = DateTime.Now;
            Trace.WriteLine(finishTime - _startTime);
            _log.Close();
        }

        private bool SetProperty<T>(ref T field, T value, [CallerMemberName] string propertyName = null)
        {
            if (Equals(field, value)) return false;
            field = value;
            OnPropertyChanged(propertyName);
            return true;
        }

        protected void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        public void CloseWindows()
        {
            _excelService.CloseExcelFile();
            _tiaService.DisposeTia();
        }
    }

    public class RelayCommand : ICommand
    {
        private readonly Action _execute;
        private readonly Func<bool> _canExecute;

        public RelayCommand(Action execute, Func<bool> canExecute = null)
        {
            _execute = execute;
            _canExecute = canExecute;
        }

        public event EventHandler CanExecuteChanged
        {
            add { CommandManager.RequerySuggested += value; }
            remove { CommandManager.RequerySuggested -= value; }
        }

        public bool CanExecute(object parameter)
        {
            return _canExecute == null || _canExecute();
        }

        public void Execute(object parameter)
        {
            _execute();
        }
    }
}