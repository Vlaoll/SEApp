using Microsoft.Extensions.Configuration;
using System;
using System.Diagnostics;

namespace seConfSW.Services
{
    public class ExcelService
    {
        private readonly IConfiguration _configuration;
        private ExcelDataReader _excelprj;
        private string _excelPath;
        private string _msg = string.Empty;

        public ExcelService(IConfiguration configuration)
        {
            _configuration = configuration;
        }

        public bool CreateExcelDB(out string message)
        {
            message = _msg;
            try
            {
                DateTime temp = DateTime.Now;
                message = temp + " :[General]Start to genetate excel DB";
                Trace.WriteLine(message);

                _excelprj = null;
                _excelprj = new ExcelDataReader();
                message = "[General]" + "##############################################################################################";
                Trace.WriteLine(message);
                if (string.IsNullOrEmpty(_excelPath))
                {
                    var filter = _configuration["Excel:Filter"] ?? "Excel |*.xlsx;*.xlsm";
                    _excelPath = _excelprj.SearchProject(filter: filter);
                }

                var mainSheetName = _configuration["Excel:MainSheetName"] ?? "Main";
                _excelprj.OpenExcelFile(_excelPath, mainSheetName: mainSheetName);

                if (!_excelprj.ReadExcelObjectData("Block", 250))
                {
                    temp = DateTime.Now;
                    message = temp + " :[General:Error] Wrong settings in excel file";
                    Trace.WriteLine(message);
                    return false;
                }
                if (!_excelprj.ReadExcelExtendedData())
                {
                    temp = DateTime.Now;
                    message = temp + " :[General:Error] Wrong settings in excel file";
                    Trace.WriteLine(message);
                    return false;
                }
                _excelprj.CloseExcelFile();

                temp = DateTime.Now;
                message = temp + " :[General]Finished to genetate excel DB";
                Trace.WriteLine(message);
                message = "[General]" + "##############################################################################################";
                Trace.WriteLine(message);

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

        public ExcelDataReader GetExcelDataReader()
        {
            return _excelprj;
        }

        public void CloseExcelFile()
        {
            try
            {
                _excelprj?.CloseExcelFile();
                _excelprj = null;
            }
            catch (Exception)
            {
                // Исключение игнорируется, как в исходном коде
            }
        }

        public string GetMessage()
        {
            return _msg;
        }
    }
}