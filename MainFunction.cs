using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WoSDataConvertor
{
    class MainFunction
    {
        static void Main(string[] args)
        {
            DataTable BeforeModifyDt = GetFileDataTables.GetBeforeModifyDt();
            string BeforeModifyFilePath = "../../Document/修改前";
            var BeforeModifyFileInfo = ReadCsvFile.ReadBeforeModifyFile(BeforeModifyFilePath);
            CsvToDataTable(BeforeModifyFileInfo);



            DirectoryInfo AfterModifyDi = new DirectoryInfo(@"../../Document/修改後");
            DirectoryInfo SummaryDi = new DirectoryInfo(@"../../Document/單一彙整檔");
            Console.WriteLine($"目前修改後的檔案資料夾路徑為：{AfterModifyDi.FullName}");
            Console.WriteLine($"目前匯出單一彙整檔的檔案資料夾路徑為：{SummaryDi.FullName}");
            Console.Read();
        }

        /// <summary>
        /// 將修改前的 CSV 經過處理後，存放到 DataTable
        /// </summary>
        /// <param name="fileInfo"></param>
        private static void CsvToDataTable(FileInfo[] fileInfo)
        {
            for (int i = 0; i < fileInfo.Length; i++)
            {
                var reader = new StreamReader(File.OpenRead(fileInfo[i].FullName));
                while (!reader.EndOfStream)
                {
                    var DataLine = reader.ReadLine();       //讀取資料列
                    var Values = DataLine.Split(',');   //用指定符號切割資料
                }
            }
        }
    }
}
