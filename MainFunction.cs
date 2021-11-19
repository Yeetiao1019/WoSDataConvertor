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
        private static DataSet BeforeModifyDs = new DataSet();
        private static DataTable BeforeModifyDt;
        private
        static void Main(string[] args)
        {
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
            int RowCount;
            for (int i = 0; i < fileInfo.Length; i++)       // 走訪修改前資料夾下的所有檔案
            {
                BeforeModifyDt = GetFileDataTables.GetBeforeModifyDt();
                RowCount = 1;
                var reader = new StreamReader(File.OpenRead(fileInfo[i].FullName));
                BeforeModifyDt.TableName = fileInfo[i].Name;
                while (!reader.EndOfStream)
                {
                    var DataLine = reader.ReadLine();       //讀取資料列
                    var Values = DataLine.Split(',');   //用指定符號切割資料
                    if (RowCount < 4 && Values.Length > 0)     // 避免讀取到格式不正確的資料列
                    {
                        RowCount++;
                    }
                    else if (RowCount >= 4 && Values.Length > 5)      // 讀取正確資料
                    {
                        for (int j = 0; j < Values[3].Split(';').Length; j++)
                        {
                            Values[3] = Values[3].Replace("; ", ";");
                            BeforeModifyDt.Rows.Add(Values[0].Replace("\"", ""),
                                Values[3].Split(';')[j].Replace("\"", ""),
                                Values[4].Replace("\"", ""),
                                Values[5].Replace("\"", ""));
                        }
                    }
                }
                BeforeModifyDs.Tables.Add(BeforeModifyDt);
            }
        }
    }
}
