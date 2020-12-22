using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.SqlClient;
using System.Data;
using DAL;
using DTO;

namespace QLGV
{
    public class BUS_khoa
    {
       DAL_Khoa  dAL_khoa = new DAL_Khoa();
        public DataTable getKhoaALL()
        {
            return dAL_khoa.getKhoa();
        }
        public bool BusThemKhoa(DTO_Khoa k)
        {
            return dAL_khoa.ThemKhoa(k);
        }
        public DataTable getTenKhoa()
        {
            return dAL_khoa.getTenKhoas();
        }
        public DataTable getBoMons(string s)
        {
            return dAL_khoa.getBoMons(s);
        }
        public DataTable getAllBoMon()
        {
            return dAL_khoa.getAllBoMon();
        }
    }
}
