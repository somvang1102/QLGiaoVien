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
    public class DAL_Khoa : DAL_con
    {
        private DAL_Khoa a;

        internal DAL_Khoa A { get => a; set => a = value; }

        public DataTable getKhoa()
        {
            SqlDataAdapter ada_khoa = new SqlDataAdapter("SELECT * FROM KHOA", conn);
            DataTable dtKhoa = new DataTable();
            ada_khoa.Fill(dtKhoa);
            return dtKhoa;
        }
        public DataTable getTenKhoas()
        {
            SqlDataAdapter ada_khoa = new SqlDataAdapter("SELECT Ma_Khoa FROM KHOA", conn);
            DataTable dtKhoa = new DataTable();
            ada_khoa.Fill(dtKhoa);
            return dtKhoa;
        }
        public DataTable getBoMons(string s)
        {
            SqlDataAdapter ada_khoa = new SqlDataAdapter("SELECT * FROM BoMon WHERE Ma_Khoa='" + s + "'", conn);
            DataTable dtKhoa = new DataTable();
            ada_khoa.Fill(dtKhoa);
            return dtKhoa;
        }
        public DataTable getAllBoMon()
        {
            SqlDataAdapter ada_khoa = new SqlDataAdapter("SELECT * FROM BoMon", conn);
            DataTable dtKhoa = new DataTable();
            ada_khoa.Fill(dtKhoa);
            return dtKhoa;
        }
        public bool ThemKhoa(DTO_Khoa k)
        {
            try
            {
                // Ket noi
                conn.Open();

                // Query string - vì mình để TV_ID là identity (giá trị tự tăng dần) nên ko cần fải insert ID
                string SQL = "INSERT INTO dbo.Khoa(Ma_Khoa, TenKhoa)VALUES('" + k.Ma_Khoa + "', N'" + k.TenKhoa1 + "')";
          

                // Command (mặc định command type = text nên chúng ta khỏi fải làm gì nhiều).
                SqlCommand cmd = new SqlCommand(SQL, conn);

                // Query và kiểm tra
                if (cmd.ExecuteNonQuery() > 0)
                    return true;

            }
            catch (Exception e)
            {

            }
            finally
            {
                // Dong ket noi
                conn.Close();
            }

            return false;

        }

    }
}
