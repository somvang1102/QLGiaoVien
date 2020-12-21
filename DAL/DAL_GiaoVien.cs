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
    public class DAL_GiaoVien:DAL_con
    {
        public List<DTO_GiaoVien> load_GV()
        {
            string query = "select * from GiaoVien";
            conn.Open();
            DataTable data = base.GetDataTable(query);
            if (data.Rows.Count == 0)
            {
                return null;
            }
            List<DTO_GiaoVien> listGV = new List<DTO_GiaoVien>();
            for (int i = 0; i < data.Rows.Count; i++)
            {
                DTO_GiaoVien gv = new DTO_GiaoVien();
                gv.Ma_GiaoVien = data.Rows[i]["Ma_GiaoVien"].ToString();
                gv.HoTen = data.Rows[i]["HoTen"].ToString();
                gv.Ma_Khoa = data.Rows[i]["Ma_Khoa"].ToString();
                gv.Ma_BoMon = data.Rows[i]["Ma_BoMon"].ToString();
                gv.GioiTinh = data.Rows[i]["GioiTinh"].ToString();
                gv.NgaySinh = DateTime.Parse(data.Rows[i]["NgaySinh"].ToString());
                gv.Email = data.Rows[i]["Email"].ToString();
                gv.ChucVu = data.Rows[i]["ChucVu"].ToString();
                gv.HocHam = data.Rows[i]["Ma_HocHam"].ToString();
                gv.HocVi = data.Rows[i]["Ma_HocVi"].ToString();
                gv.Ma_MonHoc1 = data.Rows[i]["Ma_MonHoc"].ToString();
                gv.DiaChi1= data.Rows[i]["DiaChi"].ToString();
                listGV.Add(gv);

            }
            conn.Close();
            return listGV;
        }
        public DataTable getGiaoVien(string s)
        {
            

            string query = "SELECT dbo.GiaoVien.Ma_GiaoVien,dbo.GiaoVien.HoTen,dbo.GiaoVien.NgaySinh,dbo.GiaoVien.GioiTinh,dbo.GiaoVien.Email,dbo.HocVi.TenHocVi,dbo.HocHam.TenHocHam,dbo.MonHoc.TenMonHoc,dbo.GiaoVien.ChucVu,dbo.BoMon.TenBoMon,dbo.Khoa.TenKhoa FROM dbo.GiaoVien,dbo.BoMon,dbo.HocVi,dbo.MonHoc,dbo.HocHam,dbo.Khoa WHERE dbo.GiaoVien.Ma_HocVi=dbo.HocVi.Ma_HocVi AND dbo.GiaoVien.Ma_HocHam=dbo.HocHam.Ma_HocHam AND dbo.GiaoVien.Ma_BoMon=dbo.BoMon.Ma_BoMon AND dbo.GiaoVien.Ma_MonHoc=dbo.MonHoc.Ma_MonHoc AND dbo.GiaoVien.Ma_Khoa=dbo.Khoa.Ma_Khoa AND Ma_GiaoVien='"+s+"'";

            SqlDataAdapter giaovien = new SqlDataAdapter(query, conn);
            DataTable dtKhoa = new DataTable();
            giaovien.Fill(dtKhoa);
            return dtKhoa;
        }
        public  bool Them(DTO_GiaoVien gv)
        {
            string query = "INSERT INTO dbo.GiaoVien( Ma_GiaoVien , HoTen ,NgaySinh ,GioiTinh , Email ,Ma_HocVi ,Ma_HocHam ,Ma_Khoa ,Ma_BoMon , Ma_MonHoc ,ChucVu ,DiaChi,paths )VALUES  ( '"+gv.Ma_GiaoVien+"' , N'"+gv.HoTen+"' , '"+gv.NgaySinh+"',N'"+gv.GioiTinh+"' , '"+gv.Email+"' ,  '"+gv.HocVi+"' , '"+gv.HocHam+"' , '"+gv.Ma_Khoa+"' , '"+gv.Ma_BoMon+"' , '"+gv.Ma_MonHoc1+"' , N'"+gv.ChucVu+"' , N'"+gv.DiaChi1+"','"+gv.DuongDanAnh+"' )";
            
            conn.Open();
            try
            {
                base.executeQuery(query);
                    conn.Close();
                    return true;
               
                
            }
            catch (Exception ex)
            {
                conn.Close();
                return false;
            }
        }
        public  DataTable get_MaKhoa()
        {
            conn.Open();
            string query = "SELECT Ma_Khoa FROM dbo.Khoa";
            
            DataTable data = new DataTable();
            data = base.GetDataTable(query);
           conn.Close();
            return data;
        }
        public  DataTable getMa_BoMon(string s)
        {
            conn.Open();
            string query = "SELECT Ma_BoMon FROM dbo.BoMon where Ma_Khoa ='"+s+"'";
            
            DataTable data = new DataTable();
            data = base.GetDataTable(query);
            conn.Close();
            return data;
        }
        public DataTable getTen_BoMon(string s)
        {
            conn.Open();
            string query = "SELECT TenBoMon FROM dbo.BoMon where Ma_Khoa ='" + s + "'";

            DataTable data = new DataTable();
            data = base.GetDataTable(query);
            conn.Close();
            return data;
        }
        public DataTable getHocVi()
        {
            conn.Open();
            string query = "SELECT Ma_HocVi FROM dbo.HocVi";

            DataTable data = new DataTable();
            data = base.GetDataTable(query);
            conn.Close();
            return data;
        }
        public DataTable getHocHam()
        {
            conn.Open();
            string query = "SELECT Ma_HocHam FROM dbo.HocHam";

            DataTable data = new DataTable();
            data = base.GetDataTable(query);
            conn.Close();
            return data;
        }
        public DataTable getMonHoc(string s)
        {
            conn.Open();
            string query = "SELECT Ma_MonHoc FROM MonHoc where MonHoc.Ma_BoMon='"+s+"'";

            DataTable data = new DataTable();
            data = base.GetDataTable(query);
            conn.Close();
            return data;
        }
        
        public DataTable getAnh(string s)
        {
            conn.Open();
            string query = "SELECT paths FROM dbo.GiaoVien where Ma_GiaoVien='"+s+"'";

            DataTable data = new DataTable();
            data = base.GetDataTable(query);
            conn.Close();
            return data;
        }
        public DataTable getGiaoVien2()
        {

            string query = "SELECT dbo.GiaoVien.Ma_GiaoVien,dbo.GiaoVien.HoTen,dbo.GiaoVien.NgaySinh,dbo.GiaoVien.GioiTinh,dbo.GiaoVien.Email,dbo.GiaoVien.Ma_HocVi,dbo.GiaoVien.Ma_HocHam,dbo.GiaoVien.Ma_Khoa,dbo.GiaoVien.Ma_BoMon,dbo.GiaoVien.Ma_MonHoc,dbo.GiaoVien.ChucVu,DiaChi FROM dbo.GiaoVien ";

            SqlDataAdapter giaovien = new SqlDataAdapter(query, conn);
            DataTable dtKhoa = new DataTable();
            giaovien.Fill(dtKhoa);
            return dtKhoa;
        }
        public bool themAnh(string s ,string maGiaoVien)
        {
            string query = "update GiaoVien set paths='" + s + "' where Ma_GiaoVien ='" + maGiaoVien + "' ";
            conn.Open();
            try
            {
                base.executeQuery(query);
                conn.Close();
                return true;
            }
            catch (Exception ex)
            {
                conn.Close();
                return false;
            }
        }
        // ham Xóa thong tin giao vien
        public bool xoaThongTin(string s)
        {
            string query = "DELETE dbo.GiaoVien WHERE Ma_GiaoVien ='" + s + "' ";
            conn.Open();
            try
            {
                base.executeQuery(query);
                conn.Close();
                return true;


            }
            catch (Exception ex)
            {
                conn.Close();
                return false;
            }
        }
        //Hàm sửa thông tin giáo viên
        public bool SuaThongTin(DTO_GiaoVien a)
        {
            conn.Open();           
            string SQL = "UPDATE dbo.GiaoVien SET Ma_GiaoVien='"+a.Ma_GiaoVien+"',HoTen=N'"+a.HoTen+"',NgaySinh='"+a.NgaySinh.ToString()+"',GioiTinh=N'"+a.GioiTinh+"',Email='"+a.Email+"',Ma_HocVi='"+a.HocVi+ "',Ma_HocHam='" + a.HocHam + "',Ma_Khoa='" + a.Ma_Khoa + "',Ma_BoMon='" + a.Ma_BoMon + "',Ma_MonHoc='" + a.Ma_MonHoc1 + "',ChucVu=N'" + a.ChucVu + "',DiaChi=N'" + a.DiaChi1 + "' WHERE Ma_GiaoVien='"+a.Ma_GiaoVien+"'";
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
        public DataTable timtheoma(string s)
        {
            conn.Open();
            string query=string.Format("SELECT dbo.GiaoVien.Ma_GiaoVien, dbo.GiaoVien.HoTen, dbo.GiaoVien.NgaySinh, dbo.GiaoVien.GioiTinh, dbo.GiaoVien.Email, dbo.GiaoVien.Ma_HocVi, dbo.GiaoVien.Ma_HocHam, dbo.GiaoVien.Ma_Khoa, dbo.GiaoVien.Ma_BoMon, dbo.GiaoVien.Ma_MonHoc, dbo.GiaoVien.ChucVu, DiaChi FROM dbo.GiaoVien where Ma_GiaoVien like '%"+s+"%' ");
            DataTable data = new DataTable();
            data = base.GetDataTable(query);
            conn.Close();
            return data;
        }
        public DataTable timtheoten(string s)
        {
            conn.Open();
            string query = string.Format("SELECT dbo.GiaoVien.Ma_GiaoVien, dbo.GiaoVien.HoTen, dbo.GiaoVien.NgaySinh, dbo.GiaoVien.GioiTinh, dbo.GiaoVien.Email, dbo.GiaoVien.Ma_HocVi, dbo.GiaoVien.Ma_HocHam, dbo.GiaoVien.Ma_Khoa, dbo.GiaoVien.Ma_BoMon, dbo.GiaoVien.Ma_MonHoc, dbo.GiaoVien.ChucVu, DiaChi FROM dbo.GiaoVien where HoTen like N'%" + s + "%' ");
            DataTable data = new DataTable();
            data = base.GetDataTable(query);
            conn.Close();
            return data;
        }
        // thống kê giáo viên theo khoa và bộ môn
        public DataTable ThongKe1(string s)
        {
            conn.Open();
            string query = string.Format("SELECT Ma_GiaoVien,HoTen,NgaySinh,DiaChi,ChucVu FROM dbo.GiaoVien WHERE dbo.GiaoVien.Ma_Khoa ='"+s+"'");
            DataTable data = new DataTable();
            data = base.GetDataTable(query);
            conn.Close();
            return data;
        }
        public DataTable ThongKe2(string s)
        {
            conn.Open();
            string query = string.Format("SELECT dbo.GiaoVien.Ma_GiaoVien,dbo.GiaoVien.HoTen,dbo.GiaoVien.NgaySinh,dbo.GiaoVien.DiaChi,dbo.GiaoVien.ChucVu FROM dbo.GiaoVien,dbo.BoMon WHERE dbo.GiaoVien.Ma_BoMon=dbo.BoMon.Ma_BoMon AND TenBoMon=N'"+s+"'");
            DataTable data = new DataTable();
            data = base.GetDataTable(query);
            conn.Close();
            return data;
        }
        public DataTable ThongKe3(string s,string s2)
        {
            conn.Open();
            string query = string.Format("SELECT dbo.GiaoVien.Ma_GiaoVien, dbo.GiaoVien.HoTen, dbo.GiaoVien.NgaySinh, dbo.GiaoVien.DiaChi, dbo.GiaoVien.ChucVu FROM dbo.GiaoVien WHERE Ma_Khoa= '"+s+"' AND ChucVu = N'"+s2+"'");
            DataTable data = new DataTable();
            data = base.GetDataTable(query);
            conn.Close();
            return data;
        }
        public DataTable ThongKe4( string s2)
        {
            conn.Open();
            string query = string.Format("SELECT dbo.GiaoVien.Ma_GiaoVien, dbo.GiaoVien.HoTen, dbo.GiaoVien.NgaySinh, dbo.GiaoVien.DiaChi, dbo.GiaoVien.ChucVu FROM dbo.GiaoVien WHERE ChucVu = N'" + s2 + "'");
            DataTable data = new DataTable();
            data = base.GetDataTable(query);
            conn.Close();
            return data;
        }
    }
}
