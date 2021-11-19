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
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }

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
    }
}
