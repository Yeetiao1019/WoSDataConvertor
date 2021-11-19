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
                BeforeModifyDs = ReadCsvFile.CsvToDataTable(BeforeModifyFileInfo);
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
        /// 設定領域期刊數
        /// </summary>
        private static void SetTotalJournalCount()
        {
            int FirstNaRowIndex;
            int TotalJournal;
            List<int> JournalRank = new List<int>();
            string JournalCategory;
            decimal ImpactFactor;

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
                                // 若 IF 相同
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
    }
}
