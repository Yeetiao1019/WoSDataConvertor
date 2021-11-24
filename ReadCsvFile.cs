using System;
using System.Collections.Generic;
using System.Data;
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
                if (!BeforeModifyDi.Exists)     //  若資料夾不存在，則建立
                {
                    BeforeModifyDi.Create();
                }
                Console.WriteLine($"\n目前修改前的檔案資料夾路徑為：{BeforeModifyDi.FullName}\n" +
                    $"檔案數：{BeforeModifyDi.GetFiles().Length}\n");
                return BeforeModifyDi.GetFiles();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"發生錯誤：{ex.Message}");
                return null;
            }
        }

        /// <summary>
        /// 將修改前的 CSV 經過處理後，存放到 DataTable 與 DataSet
        /// </summary>
        /// <param name="fileInfo"></param>
        public static DataSet CsvToDataTable(FileInfo[] fileInfo)
        {
            DataSet Ds = new DataSet();
            DataTable Dt = new DataTable();
            int RowCount;
            string Category;
            string JournalRank;
            string TotalJournal;
            decimal ImpactFactor;
            string Quartile;
            try
            {
                for (int i = 0; i < fileInfo.Length; i++)       // 走訪修改前資料夾下的所有檔案
                {
                    Dt = GetFileDataTables.GetBeforeModifyDt();
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
                        else if (RowCount >= 4 && Values.Length > 5 && Values[5].Trim().Length != 0)      // 讀取正確資料
                        {
                            for (int j = 0; j < Values[3].Split(';').Length; j++)
                            {
                                Category = Values[3];
                                // 收錄分類為 ESCI 或 AHCI 或 IF 值為 n/a，則領域期刊數、期刊排名、分位數為 n/a
                                if (Values[3].Split(';')[j].Contains("ESCI") || Values[3].Split(';')[j].Contains("AHCI") || Values[5].Contains("n/a"))
                                {
                                    JournalRank = "n/a";
                                    TotalJournal = "n/a";
                                    Quartile = "n/a";
                                    ImpactFactor = 0.00001m;
                                }
                                else
                                {
                                    JournalRank = "1";
                                    TotalJournal = "";
                                    Quartile = Values[4].Replace("\"", "");
                                    ImpactFactor = Convert.ToDecimal(Values[5].Replace("\"", ""));
                                }
                                Category = Category.Replace("; ", ";").Replace("，", ", ");     //將全形逗號還原回來
                                Dt.Rows.Add(Values[0].Replace("\"", ""),
                                    Category.Split(';')[j].Replace("\"", ""),
                                    Quartile,
                                    ImpactFactor,
                                    JournalRank,
                                    TotalJournal);
                            }
                        }
                    }
                    Dt = Dt.AsEnumerable()      // 以Category降冪排序
                        .OrderBy(r => r.Field<string>("Category"))
                        //.OrderByDescending(r => r.Field<decimal>("JIF"))
                        .CopyToDataTable();
                    //Dt = Dt.AsEnumerable()      // 以 IF 值降冪排序

                    //    .CopyToDataTable();
                    Dt.TableName = fileInfo[i].Name.Replace(".csv", "");
                    Ds.Tables.Add(Dt);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("發生錯誤：讀取 csv 發生錯誤");
            }

            return Ds;
        }
    }
}
