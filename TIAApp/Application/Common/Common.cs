using Siemens.Engineering.SW.Tags;
using System.Collections.Generic;
using System.IO;
using System.Text;
using seConfSW.Domain.Models;

namespace seConfSW
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
