using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.Data.SqlClient;
using DTO;
namespace DAL
{
    public class DAL_HocHam:DAL_con
    {
        public static DataTable dt;
        public DataTable thongkeHocHam(string truyvan)
        {
            SqlDataAdapter ada_khoa = new SqlDataAdapter(truyvan, conn);
            DataTable dtKhoa = new DataTable();
            ada_khoa.Fill(dtKhoa);
            return dtKhoa;
        }

    }
}
