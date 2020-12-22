using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using DAL;
using DTO;
namespace QLGV
{
    public class Bus_DangNhap
    {
       
        DAL_DangNhap _DangNhap = new DAL_DangNhap();
        public bool DangNhap(string taiKhoan,string matKhau)
        {
            return _DangNhap.DangNhap(taiKhoan, matKhau);
        }
        public bool doiMatKhau(string tenND, string matKhau)
        {
            return _DangNhap.layTenDangNhap(tenND, matKhau);
        }
        public DataTable getAccount(string s)
        {
            return _DangNhap.getAccount("SELECT dbo.Account.UserName,dbo.Account.DisplayName,dbo.Account.Type,dbo.GiaoVien.HoTen,dbo.GiaoVien.Email FROM dbo.Account,dbo.GiaoVien WHERE Account.Ma_GiaoVien=dbo.GiaoVien.Ma_GiaoVien and dbo.Account.UserName='" + s + "'");
        }
        public List<DTO_DangNhap> getTaiKhoan()
        {
            return _DangNhap.getTaiKhoan();
        }
    }
}
