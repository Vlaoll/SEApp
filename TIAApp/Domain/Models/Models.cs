

using Siemens.Engineering.SW.Tags;
using System.Collections.Generic;
using System.Text;

namespace seConfSW.Domain.Models
{
    public class dataPLC
    {
        public string namePLC { get; set; }
        public List<dataEq> Equipment { get; set; } = new List<dataEq>();
        public List<dataBlock> instanceDB { get; set; } = new List<dataBlock>();
        public List<dataFunction> dataFC { get; set; } = new List<dataFunction>();
        public List<dataSupportBD> dataSupportBD { get; set; } = new List<dataSupportBD>();
        public List<userConstant> userConstant { get; set; } = new List<userConstant>();
    }

    public class dataBlock
    {
        public string name { get; set; }
        public string comment { get; set; }
        public string area { get; set; }
        public string instanceOfName { get; set; }
        public int number { get; set; }
        public string group { get; set; }
        public string nameFC { get; set; }
        public string typeEq { get; set; }
        public string nameEq { get; set; }
        public List<string> variant { get; set; } = new List<string>();
        public List<excelData> excelData { get; set; } = new List<excelData>();
    }

    public class dataExtSupportBlock
    {
        public string name { get; set; }
        public string instanceOfName { get; set; }
        public int number { get; set; }
        public List<string> variant { get; set; } = new List<string>();
    }

    public class dataFunction
    {
        public string name { get; set; }
        public int number { get; set; }
        public string group { get; set; }
        public StringBuilder code { get; set; } = new StringBuilder();
        public List<dataExtFCValue> dataExtFCValue { get; set; } = new List<dataExtFCValue>();
    }

    public class dataEq
    {
        public string typeEq { get; set; }
        public bool isExtended { get; set; }
        public List<dataLibrary> FB { get; set; } = new List<dataLibrary>();
        public List<dataTag> dataTag { get; set; } = new List<dataTag>();
        public List<dataParameter> dataParameter { get; set; } = new List<dataParameter>();
        public List<dataExtSupportBlock> dataExtSupportBlock { get; set; } = new List<dataExtSupportBlock>();
        public List<userConstant> dataConstant { get; set; } = new List<userConstant>();
        public List<dataSupportBD> dataSupportBD { get; set; } = new List<dataSupportBD>();
        public List<dataDataBlockValue> dataDataBlockValue { get; set; } = new List<dataDataBlockValue>();
    }

    public class dataTag
    {
        public string name { get; set; }
        public string link { get; set; }
        public string type { get; set; }
        public string adress { get; set; }
        public string table { get; set; }
        public string comment { get; set; }
        public List<string> variant { get; set; } = new List<string>();
    }

    public class dataParameter
    {
        public string name { get; set; }
        public string type { get; set; }
        public string link { get; set; }
    }

    public class excelData
    {
        public string name { get; set; }
        public int column { get; set; }
        public string value { get; set; }
    }

    public class userConstant
    {
        public string name { get; set; }
        public string type { get; set; }
        public string value { get; set; }
        public string table { get; set; }
    }

    public class dataSupportBD
    {
        public string name { get; set; }
        public int number { get; set; }
        public string group { get; set; }
        public string type { get; set; }
        public string path { get; set; }
        public bool isType { get; set; }
        public bool isRetain { get; set; }
        public bool isOptimazed { get; set; }
    }

    public class dataLibrary
    {
        public string name { get; set; }
        public string path { get; set; }
        public string group { get; set; }
        public bool isType { get; set; }
    }

    public class dataImportTag
    {
        public string name { get; set; }
        public string table { get; set; }
        public StringBuilder code { get; set; } = new StringBuilder();
        public string fileName { get; set; }
        public int ID { get; set; }
    }

    public class dataExcistTag
    {
        public string table { get; set; }
        public List<PlcTag> tags { get; set; } = new List<PlcTag>();
    }

    public class dataDataBlockValue
    {
        public string name { get; set; }
        public string type { get; set; }
        public string DB { get; set; }
    }

    public class dataExtFCValue
    {
        public string name { get; set; }
        public string type { get; set; }
        public string IO { get; set; }
        public string Comments { get; set; }
    }
}
