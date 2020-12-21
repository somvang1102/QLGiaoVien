using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DTO
{
    public class DTO_DangNhap
    {
        private string tenNguoiDung;
        private string tenTaiKhoa;
        private string matKhau;
        private int loaiTaiKhoan;
        private string maND;

        public string TenNguoiDung { get => tenNguoiDung; set => tenNguoiDung = value; }
        public string TenTaiKhoa { get => tenTaiKhoa; set => tenTaiKhoa = value; }
        public string MatKhau { get => matKhau; set => matKhau = value; }
        public int LoaiTaiKhoan { get => loaiTaiKhoan; set => loaiTaiKhoan = value; }
        public string MaND { get => maND; set => maND = value; }
    }
}
