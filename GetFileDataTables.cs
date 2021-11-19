using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WoSDataConvertor
{
    public static class GetFileDataTables
    {
        public static DataTable GetBeforeModifyDt()
        {
            DataTable Dt = new DataTable();
            Dt.Columns.Add("Journal Name", typeof(string));
            Dt.Columns.Add("Category", typeof(string));
            Dt.Columns.Add("JIF Quartile", typeof(string));
            Dt.Columns.Add("JIF", typeof(decimal));
            Dt.Columns.Add("期刊排名(Journal Rank)", typeof(string));
            Dt.Columns.Add("領域期刊數(Total Journal)", typeof(string));

            return Dt;
        }
    }
}
