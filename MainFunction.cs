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
        static void Main(string[] args)
        {
            string BeforeModifyFilePath = "../../Document/修改前";
            string AfterModifyFilePath = "../../Document/修改後";
            string SummaryPath = "../../Document/單一彙整檔";
            try
            {
                var BeforeModifyFileInfo = ReadCsvFile.ReadBeforeModifyFile(BeforeModifyFilePath);
                BeforeModifyDs = ReadCsvFile.CsvToDataTable(BeforeModifyFileInfo);
                SetTotalJournalCountAndRank();

                DirectoryInfo AfterModifyDi = new DirectoryInfo(AfterModifyFilePath);
                if (!AfterModifyDi.Exists)     //  若資料夾不存在，則建立
                {
                    AfterModifyDi.Create();
                }
                Console.WriteLine($"目前修改後的檔案資料夾路徑為：{AfterModifyDi.FullName}");
                for (int i = 0; i < BeforeModifyDs.Tables.Count; i++)
                {
                    var Dt = BeforeModifyDs.Tables[i];
                    Dt = ReplaceToNA(Dt);
                    GenerateExcel.ExportExcelFile(Dt, Dt.TableName, AfterModifyFilePath);
                }
                Console.WriteLine($"檔案輸出完成，共 {GenerateExcel.FileCount} 個檔案。\n");

                DirectoryInfo SummaryDi = new DirectoryInfo(SummaryPath);
                if (!SummaryDi.Exists)     //  若資料夾不存在，則建立
                {
                    SummaryDi.Create();
                }
                Console.WriteLine($"目前匯出單一彙整檔的檔案資料夾路徑為：{SummaryDi.FullName}");
                GenerateExcel.ExportOneExcelFile(BeforeModifyDs, SummaryPath);                
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }

            Console.Read();
        }

        /// <summary>
        /// 設定領域期刊數與名次
        /// </summary>
        private static void SetTotalJournalCountAndRank()
        {
            Console.WriteLine("處理排序中...");
            int TotalJournal;
            List<int> JournalRank = new List<int>();

            for (int i = 0; i < BeforeModifyDs.Tables.Count; i++)
            {
                var Dt = BeforeModifyDs.Tables[i];

                var JournalCategoryRow = Dt.AsEnumerable()     // 取得 Dt 內的所有 Category
                    .GroupBy(r => r.Field<string>("Category"))
                    .ToList();

                for (int j = 0; j < JournalCategoryRow.Count; j++)
                {
                    JournalRank.Clear();

                    // 避免讀到全為 n/a 的收錄分類
                    if (Dt.AsEnumerable().Where(r => r.Field<string>("Category") == JournalCategoryRow[j].Key &&
                    r.Field<string>("領域期刊數(Total Journal)") != "n/a").ToList().Count != 0)
                    {
                        TotalJournal = Dt.AsEnumerable()        // 取得該 Category 的領域期刊數（不包含 n/a）
                    .Where(r => r.Field<string>("Category") == JournalCategoryRow[j].Key &&
                    r.Field<string>("領域期刊數(Total Journal)") != "n/a")
                    .CopyToDataTable().Rows.Count;

                        var AllDataRow = Dt.AsEnumerable()     // 取得該 Catetory 的所有 DataRow（不包含n/a）
                            .Where(r => r.Field<string>("Category") == JournalCategoryRow[j].Key &&
                            r.Field<string>("領域期刊數(Total Journal)") != "n/a");

                        // 設定領域期刊數的值
                        AllDataRow.ToList().ForEach(r => r.SetField<string>("領域期刊數(Total Journal)", TotalJournal.ToString()));

                        for (int k = 0; k < AllDataRow.ToList().Count; k++)
                        {
                            JournalRank.Add(2);
                        }

                        // 排名排序
                        for (int k = 1; k < AllDataRow.ToList().Count; k++)
                        {
                            for (int m = 0; m < k; m++)
                            {
                                if (Convert.ToDecimal(AllDataRow.ToList()[m]["JIF"]) < Convert.ToDecimal(AllDataRow.ToList()[k]["JIF"]))
                                {
                                    AllDataRow.ToList()[m]["期刊排名(Journal Rank)"] = JournalRank[m]++;
                                }
                                // 若 IF 相同，則名次相同
                                else if (Convert.ToDecimal(AllDataRow.ToList()[m]["JIF"]) == Convert.ToDecimal(AllDataRow.ToList()[k]["JIF"]))
                                {
                                    AllDataRow.ToList()[m]["期刊排名(Journal Rank)"] = JournalRank[m] - 1;
                                }
                                else
                                {
                                    AllDataRow.ToList()[k]["期刊排名(Journal Rank)"] = JournalRank[k]++;
                                }
                            }
                        }
                    }
                }
            }
        }

        /// <summary>
        /// 替換回 n/a
        /// </summary>
        /// <param name="dt"></param>
        public static DataTable ReplaceToNA(DataTable dt)
        {
            DataTable AfterChangeTypeDt = new DataTable();
            AfterChangeTypeDt = dt.Clone();
            AfterChangeTypeDt.Columns["JIF"].DataType = typeof(string);
            foreach (DataRow r in dt.Rows)
            {
                AfterChangeTypeDt.Rows.Add(r.ItemArray);
            }

            for (int i = 0; i < AfterChangeTypeDt.Rows.Count; i++)
            {
                if (AfterChangeTypeDt.Rows[i]["JIF"].ToString() == "0.00001")
                {
                    AfterChangeTypeDt.Rows[i]["JIF"] = "n/a";
                }
            }

            return AfterChangeTypeDt;
        }
    }
}
