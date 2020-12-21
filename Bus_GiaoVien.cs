using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DTO;
using DAL;
using System.Data;
using System.Data.SqlClient;
namespace QLGV
{
    
    public class Bus_GiaoVien
    {
        DAL_GiaoVien dAL = new DAL_GiaoVien();
        public  List<DTO_GiaoVien> load_GV()
        {
            return dAL.load_GV();
        }
        public  bool Them(DTO_GiaoVien gv)
        {
            return dAL.Them(gv);
        }
        public  DataTable getMa_Khoa()
        {
            return dAL.get_MaKhoa();
        }
        public  DataTable getMa_BoMon(string s)
        {
            return dAL.getMa_BoMon(s);
        }
        public DataTable getTen_BoMon(string s)
        {
            return dAL.getTen_BoMon(s);
        }
        public DataTable getGiaoVien(string s)
        {
            return dAL.getGiaoVien(s);
        }
        public DataTable getGiaoVien2()
        {
            return dAL.getGiaoVien2();
        }
        public DataTable getHocHam()
        {
            return dAL.getHocHam();

        }
        public DataTable getHocVi()
        {
            return dAL.getHocVi();

        }
        public DataTable getPath(string s)
        {
            return dAL.getAnh(s);
        }
        public DataTable getMonHoc(string s)
        {
            return dAL.getMonHoc(s);    
        }
        public bool ThemAnh(string s,string maGiaoVien)
        {
            return dAL.themAnh(s,maGiaoVien);
        }
        public bool xoa(string s)
        {
            return dAL.xoaThongTin(s);
        }
        public bool sua(DTO_GiaoVien a)
        {
            return dAL.SuaThongTin(a);
        }
        public DataTable timtheoma(string s)
        {
            return dAL.timtheoma(s);
        }
        public DataTable timtheoten(string s)
        {
            return dAL.timtheoten(s);
        }
        public DataTable ThongKe1(string s)
        {
            return dAL.ThongKe1(s);
        }
        public DataTable ThongKe2(string s)
        {
            return dAL.ThongKe2(s);
        }
        public DataTable ThongKe3(string s,string s2)
        {
            return dAL.ThongKe3(s,s2);
        }
        public DataTable ThongKe4(string s)
        {
            return dAL.ThongKe4(s);

            }

    }
}
