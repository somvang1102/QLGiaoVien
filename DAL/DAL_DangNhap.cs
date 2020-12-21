using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DTO;
using System.Data;
using System.Data.SqlClient;


namespace DAL
{
    public class DAL_DangNhap:DAL_con
    {
        
        public static DataTable dt;
        DTO_DangNhap a = new DTO_DangNhap();

        public DTO_DangNhap A { get => a; set => a = value; }

        public DataTable getAccount(string truyvan)
        {
            SqlDataAdapter ada_khoa = new SqlDataAdapter(truyvan, conn);
            DataTable dtKhoa = new DataTable();
            ada_khoa.Fill(dtKhoa);
            return dtKhoa;
        }
        
        public List<DTO_DangNhap> getTaiKhoan()
        {
            string truyvan = "SELECT * FROM Account";
            List<DTO_DangNhap> l = new List<DTO_DangNhap>();
            DTO_DangNhap a = new DTO_DangNhap();
            dt = getAccount(truyvan);
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                a.LoaiTaiKhoan= Convert.ToInt32(dt.Rows[i]["Type"].ToString());
                a.TenNguoiDung= dt.Rows[i]["DisPlayName"].ToString();
                a.TenTaiKhoa= dt.Rows[i]["UserName"].ToString();
                a.MatKhau= dt.Rows[i]["PassWord"].ToString();
                l.Add(a);
            }
            return l; 
        }
        public bool DangNhap(string taiKhoan,string matKhau)
        {
            conn.Open();
            
            string SQL = "SELECT * FROM dbo.Account WHERE UserName='"+taiKhoan+"'AND PassWord='"+matKhau+"'";
            SqlCommand cmd = new SqlCommand(SQL, conn);
            SqlDataReader s = cmd.ExecuteReader();
            if (s.Read() == true)
            {
                conn.Close();
                return true;
            }

            else
            {
                conn.Close();
                return false;
            }

        }
        
        public bool layTenDangNhap(string tenDangNhap, string matKhau)
        {
            conn.Open();

            //Query string - vì mình để TV_ID là identity (giá trị tự tăng dần) nên ko cần fải insert ID
            string SQL = "UPDATE dbo.Account SET PassWord='" + matKhau + "'WHERE UserName='" + tenDangNhap + "'";


            //Command (mặc định command type = text nên chúng ta khỏi fải làm gì nhiều).
            try
            {
                SqlCommand cmd = new SqlCommand(SQL, conn);

                // Query và kiểm tra
                cmd.ExecuteNonQuery();
                conn.Close();
                return true;
            }

          

            catch (Exception ex)
            {
                conn.Close();
                return false;
            }
        }
        //hàm lấy về thông tin tài khoản 
        
    }
}
