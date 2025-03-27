// Ignore Spelling: Conf

using Siemens.Engineering.SW;
using System;
using System.Collections.Generic;

namespace seConfSW.Services
{
    public interface ITiaService
    {       
        event EventHandler<string> MessageUpdated; 
        bool AddValueToDataBlock(PlcSoftware plcSoftware, Domain.Models.dataPLC dataPLC);
        bool CreateCommonUserConstants(PlcSoftware plcSoftware, Domain.Models.dataPLC dataPLC);
        bool CreateEqConstants(PlcSoftware plcSoftware, Domain.Models.dataPLC dataPLC);
        bool CreateTagsFromFile(PlcSoftware plcSoftware, Domain.Models.dataPLC dataPLC);
        bool UpdatePrjLibraryFromGlobal(PlcSoftware plcSoftware, Domain.Models.dataPLC dataPLC, string libraryPath);
        bool UpdateSupportBlocks(PlcSoftware plcSoftware, Domain.Models.dataPLC dataPLC, string libraryPath);
        bool UpdateTypeBlocks(PlcSoftware plcSoftware, Domain.Models.dataPLC dataPLC, string libraryPath);
        bool CreateInstanceBlocks(PlcSoftware plcSoftware, Domain.Models.dataPLC dataPLC);
        bool CreateTemplateFCFromExcel(PlcSoftware plcSoftware, Domain.Models.dataPLC dataPLC);
        bool EditFCFromExcelCallAllBlocks(PlcSoftware plcSoftware, Domain.Models.dataPLC dataPLC, List<Domain.Models.dataPLC> listDataPLC, bool closeProject, bool saveProject, bool compileProject);
        void DisposeTia();
    }
}