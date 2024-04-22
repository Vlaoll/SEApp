using Siemens.Engineering.SW.Tags;
using System.Collections.Generic;
using System.IO;
using System.Text;


namespace TIAApp
{
    static class Common
    {
        public static string ModifyString(string excString, List<excelData> excelData)
        {
            if (excelData.Count > 0)
            {
                foreach (var item in excelData)
                {
                    if (excString.Contains(item.name))
                    {
                        excString = excString.Replace(item.name, item.value).Trim();
                    }
                }
            }
            return excString;
        }
        public static string ModifyString(string excString, List<excelData> excelData, string preFix)
        {
            if (excelData.Count > 0)
            {
                foreach (var item in excelData)
                {
                    if (excString.Contains(preFix + item.name))
                    {
                        excString = excString.Replace(preFix + item.name, item.value).Trim();
                    }
                }
            }
            return excString;
        }
        public static void SafeDictionaryAdd(Dictionary<string, object> dict, string key, object view)
        {
            if (!dict.ContainsKey(key))
            {
                dict.Add(key, view);
            }
            else
            {
                dict[key] = view;
            }
        }
        public static void SafeDictionaryAdd(Dictionary<string, object> dict, string key)
        {
            if (!dict.ContainsKey(key))
            {
                dict.Add(key, null);
            }
        }

        public static void CreateNewFolder(string path)
        {
            try
            {
                if (Directory.Exists(path)) Directory.Delete(path, true);
                Directory.CreateDirectory(path);
            }
            catch (System.Exception)
            {                
            }
           
        }

    }
}
public struct dataPLC
{
    public string namePLC;
    public List<dataEq> Equipment;
    public List<dataBlock> instanceDB;    
    public List<dataFunction> dataFC;
    public List<dataSupportBD> dataSupportBD;
    public List<userConstant> userConstant;
    
}
public struct dataBlock
{
    public string name;
    public string comment;
    public string area;
    public string instanceOfName;
    public int number;
    public string group;
    public string nameFC;
    public string typeEq;
    public string nameEq;
    public List<string> variant;
    public List<excelData> excelData;
}
public struct dataExtSupportBlock
{
    public string name;    
    public string instanceOfName;
    public int number;
    public List<string> variant;
}
public struct dataFunction
{
    public string name;
    public int number;
    public string group;
    public StringBuilder code;
    public List<dataExtFCValue> dataExtFCValue;
}
public struct dataEq
{
    public string typeEq;
    public bool isExtended;
    public List<dataLibrary> FB;
    public List<dataTag> dataTag;
    public List<dataParameter> dataParameter;
    public List<dataExtSupportBlock> dataExtSupportBlock;
    public List<userConstant> dataConstant;
    public List<dataSupportBD> dataSupportBD;
    public List<dataDataBlockValue> dataDataBlockValue;
}
public struct dataTag
{
    public string name ;
    public string link ;
    public string type ;
    public string adress;
    public string table;
    public string comment;
    public List<string> variant;
}
public struct dataParameter
{
    public string name;
    public string type;
    public string link;
}
public struct excelData
{
    public string name;
    public int column;
    public string value;
}
public struct userConstant
{
    public string name;
    public string type;
    public string value;
    public string table;
}
public struct dataSupportBD
{
    public string name;
    public int number;
    public string group;
    public string type;
    public string path;
    public bool isType;
    public bool isRetain;
    public bool isOptimazed;
}
public struct dataLibrary
{
    public string name;
    public string path;
    public string group;
    public bool isType;
}
public struct dataImportTag
{
    public string name;
    public string table;
    public StringBuilder code;
    public string path;
    public int ID;    
}
public struct dataExcistTag
{
    public string table;
    public List<PlcTag> tags; 
}
public struct dataDataBlockValue
{
    public string name;
    public string type;
    public string DB;   
}
public struct dataExtFCValue
{
    public string name;
    public string type;
    public string IO;
    public string Comments;
}