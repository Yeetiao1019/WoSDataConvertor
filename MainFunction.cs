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
            try
            {
                var BeforeModifyFileInfo = ReadCsvFile.ReadBeforeModifyFile(BeforeModifyFilePath);
                CsvToDataTable(BeforeModifyFileInfo);
                SetTotalJournalCount();

                DirectoryInfo AfterModifyDi = new DirectoryInfo(@"../../Document/修改後");
                DirectoryInfo SummaryDi = new DirectoryInfo(@"../../Document/單一彙整檔");
                Console.WriteLine($"目前修改後的檔案資料夾路徑為：{AfterModifyDi.FullName}");
                Console.WriteLine($"目前匯出單一彙整檔的檔案資料夾路徑為：{SummaryDi.FullName}");
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }

            Console.Read();
        }

        /// <summary>
        /// 將修改前的 CSV 經過處理後，存放到 DataTable
        /// </summary>
        /// <param name="fileInfo"></param>
        private static void CsvToDataTable(FileInfo[] fileInfo)
        {
            int RowCount;
            string Category;
            string JournalRank;
            string TotalJournal;
            string ImpactFactor;
            string Quartile;
            try
            {
                for (int i = 0; i < fileInfo.Length; i++)       // 走訪修改前資料夾下的所有檔案
                {
                    BeforeModifyDt = GetFileDataTables.GetBeforeModifyDt();
                    RowCount = 1;
                    var reader = new StreamReader(File.OpenRead(fileInfo[i].FullName));
                    while (!reader.EndOfStream)
                    {
                        var DataLine = reader.ReadLine().Replace(", ", "，");       //讀取資料列，並先取代 "逗號+空白" 避免因 Category 的名稱導致資料切分錯誤
                        var Values = DataLine.Split(',');   //用指定符號切割資料
                        if (RowCount < 4 && Values.Length > 0)     // 避免讀取到格式不正確的資料列
                        {
                            RowCount++;
                        }
                        else if (RowCount >= 4 && Values.Length > 5)      // 讀取正確資料
                        {
                            for (int j = 0; j < Values[3].Split(';').Length; j++)
                            {
                                Category = Values[3];
                                // IF值為 n/a，則領域期刊數、期刊排名、分位數為 n/a
                                if (Values[5].Contains("n/a"))
                                {
                                    JournalRank = "n/a";
                                    TotalJournal = "n/a";
                                    Quartile = "n/a";
                                    ImpactFactor = "n/a";
                                }
                                else
                                {
                                    JournalRank = "";
                                    TotalJournal = "";
                                    Quartile = Values[4].Replace("\"", "");
                                    ImpactFactor = Values[5].Replace("\"", "");
                                }
                                Category = Category.Replace("; ", ";").Replace("，", ", ");     //將全形逗號還原回來
                                BeforeModifyDt.Rows.Add(Values[0].Replace("\"", ""),
                                    Category.Split(';')[j].Replace("\"", ""),
                                    Quartile,
                                    ImpactFactor,
                                    JournalRank,
                                    TotalJournal);
                            }
                        }
                    }
                    BeforeModifyDt = BeforeModifyDt.AsEnumerable()
                        .OrderBy(r => r.Field<string>("JIF"))
                        .CopyToDataTable();
                    BeforeModifyDt.TableName = fileInfo[i].Name.Replace(".csv", "");
                    BeforeModifyDs.Tables.Add(BeforeModifyDt);
                }
            }
            catch (Exception)
            {
                Console.WriteLine("讀取 csv 發生錯誤");
            }
        }

        /// <summary>
        /// 設定領域期刊數
        /// </summary>
        private static void SetTotalJournalCount()
        {
            for (int i = 0; i < BeforeModifyDs.Tables.Count; i++)
            {
                var Dt = BeforeModifyDs.Tables[i];
                var JournalCategory = Dt.AsEnumerable()     // 取得 Dt 內的所有 Category
                    .GroupBy(r => r.Field<string>("Category"))
                    .ToList();

                for (int j = 0; j < JournalCategory.Count; j++)
                {

                }

            }
        }
    }
}
