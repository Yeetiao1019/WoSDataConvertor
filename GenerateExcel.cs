using NPOI.SS.UserModel;
using NPOI.XSSF.Streaming;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WoSDataConvertor
{
    public static class GenerateExcel
    {
        public static int FileCount { get; set; } = 0;
        public static void ExportExcelFile(DataTable dt, string category, string path)
        {
            try
            {
                using (FileStream stream = new FileStream($@"{path}/{category}.xlsx", FileMode.Create, FileAccess.Write))
                {
                    SXSSFWorkbook Workbook = new SXSSFWorkbook();
                    ISheet Sheet;
                    Sheet = Workbook.CreateSheet(category);
                    Sheet = DataTableToSheet(0, Sheet, dt, true);
                    Workbook.Write(stream);
                    FileCount++;
                    Console.WriteLine($"已輸出成功檔案：{FileCount}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }

        /// <summary>
        /// 匯出單檔 Excel
        /// </summary>
        /// <param name="ds"></param>
        /// <param name="path"></param>
        public static void ExportOneExcelFile(DataSet ds, string path)
        {
            string Category;
            try
            {
                Category = "彙整總檔";
                using (FileStream stream = new FileStream($@"{path}/{Category}.xlsx", FileMode.Create, FileAccess.Write))
                {
                    SXSSFWorkbook Workbook = new SXSSFWorkbook();
                    ISheet Sheet;
                    Sheet = Workbook.CreateSheet(Category);
                    Sheet = DataSetToSheet(0, Sheet, ds, true);
                    Workbook.Write(stream);
                    Console.WriteLine("已完成檔案輸出");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"發生錯誤：{ex.Message}");
            }
        }

        /// <summary>
        /// 傳入 DataTable，將 DataTable 輸出為 Excel Sheet。
        /// </summary>
        /// <param name="startRowIndex">Excel 要從第幾列開始列出資料</param>
        /// <param name="sheet"></param>
        /// <param name="dt">傳入 DataTable</param>
        /// <param name="isShowColumn">Excel 是否要列出 DataTable 欄位名稱</param>
        /// <returns></returns>
        private static ISheet DataTableToSheet(int startRowIndex, ISheet sheet, DataTable dt, bool isShowColumn)
        {
            if (dt != null && isShowColumn)
            {
                IRow row = sheet.CreateRow(startRowIndex);
                for (int i = 0; i < dt.Columns.Count; i++)
                {
                    row.CreateCell(i).SetCellValue(dt.Columns[i].ToString());
                }

                startRowIndex++;
            }

            if (dt.Rows.Count > 0)
            {
                for (int j = 0; j < dt.Rows.Count; j++)
                {
                    IRow row = sheet.CreateRow(startRowIndex);
                    for (int k = 0; k < dt.Columns.Count; k++)
                    {
                        switch (k)
                        {
                            default:
                                row.CreateCell(k).SetCellValue(dt.Rows[j][dt.Columns[k]?.ToString() ?? ""]?.ToString() ?? "");
                                break;
                        }
                    }

                    startRowIndex++;
                }
            }

            return sheet;
        }

        /// <summary>
        /// 傳入 DataSet，將 DataSet 內的所有 DataTable 輸出為 Excel Sheet。
        /// </summary>
        /// <param name="startRowIndex">Excel 要從第幾列開始列出資料</param>
        /// <param name="sheet"></param>
        /// <param name="ds">傳入 DataSet</param>
        /// <param name="isShowColumn">Excel 是否要列出 DataTable 欄位名稱</param>
        /// <returns></returns>
        private static ISheet DataSetToSheet(int startRowIndex, ISheet sheet, DataSet ds, bool isShowColumn)
        {
            if (ds != null && isShowColumn)
            {
                IRow row = sheet.CreateRow(startRowIndex);
                for (int j = 0; j < ds.Tables[0].Columns.Count; j++)
                {
                    row.CreateCell(j).SetCellValue(ds.Tables[0].Columns[j].ToString());
                }

                startRowIndex++;
            }

            for (int i = 0; i < ds.Tables.Count; i++)
            {
                var Dt = MainFunction.ReplaceToNA(ds.Tables[i]);

                if (Dt.Rows.Count > 0)
                {
                    for (int j = 0; j < Dt.Rows.Count; j++)
                    {
                        IRow row = sheet.CreateRow(startRowIndex);
                        for (int k = 0; k < Dt.Columns.Count; k++)
                        {
                            switch (k)
                            {
                                default:
                                    row.CreateCell(k).SetCellValue(Dt.Rows[j][Dt.Columns[k]?.ToString() ?? ""]?.ToString() ?? "");
                                    break;
                            }
                        }
                        startRowIndex++;
                    }
                }
            }

            return sheet;
        }
    }
}
