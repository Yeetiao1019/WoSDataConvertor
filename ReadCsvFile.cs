using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WoSDataConvertor
{
    public static class ReadCsvFile
    {
        /// <summary>
        /// 讀取修改前的檔案，並回傳 FileInfo[]
        /// </summary>
        public static FileInfo[] ReadBeforeModifyFile(string path)
        {
            try
            {
                DirectoryInfo BeforeModifyDi = new DirectoryInfo($@"{path}");
                Console.WriteLine($"\n目前修改前的檔案資料夾路徑為：{BeforeModifyDi.FullName}");
                return BeforeModifyDi.GetFiles();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                return null;
            }
        }
    }
}
