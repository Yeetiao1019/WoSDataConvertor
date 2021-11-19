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
        private static DataTable BeforeModifyDt;
        static void Main(string[] args)
        {
            BeforeModifyDt = GetFileDataTables.GetBeforeModifyDt();
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
            bool IsTitleRow;
            for (int i = 0; i < fileInfo.Length; i++)
            {
                IsTitleRow = false;
                var reader = new StreamReader(File.OpenRead(fileInfo[i].FullName));
                while (!reader.EndOfStream)
                {
                    var DataLine = reader.ReadLine();       //讀取資料列
                    var Values = DataLine.Split(',');   //用指定符號切割資料
                    if (Values.Length == 6 && IsTitleRow == false && Values[5].Length > 0)     // 避免讀取到格式不正確的資料列
                    {
                        IsTitleRow = true;
                    }
                    else if (Values.Length == 6 && IsTitleRow == true && Values[5].Length > 0)      // 讀取正確資料
                    {
                        for (int j = 0; j < Values[3].Split(';').Length; j++)
                        {
                            BeforeModifyDt.Rows.Add(Values[0], Values[3].Split(';')[j], Values[4], Values[5]);
                        }
                    }
                }
            }
        }
    }
}
