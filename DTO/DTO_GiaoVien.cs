using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DTO
{
    public class DTO_GiaoVien
    {
        private string ma_GiaoVien;
        private string hoTen;
        private DateTime ngaySinh;
        private string email;
        private string ma_Khoa;
        private string ma_BoMon;       
        private string gioiTinh;
        private string hocVi;
        private string hocHam;
        private string chucVu;
        private string duongDanAnh;
        private string Ma_MonHoc;
        private string DiaChi;

        
        public string Ma_GiaoVien { get => ma_GiaoVien; set => ma_GiaoVien = value; }
        public string HoTen { get => hoTen; set => hoTen = value; }

        public string Email { get => email; set => email = value; }
        public string Ma_Khoa { get => ma_Khoa; set => ma_Khoa = value; }
        public string Ma_BoMon { get => ma_BoMon; set => ma_BoMon = value; }
        public string GioiTinh { get => gioiTinh; set => gioiTinh = value; }
        public DateTime NgaySinh { get => ngaySinh; set => ngaySinh = value; }
        public string HocVi { get => hocVi; set => hocVi = value; }
        public string HocHam { get => hocHam; set => hocHam = value; }
        public string ChucVu { get => chucVu; set => chucVu = value; }
        public string DuongDanAnh { get => duongDanAnh; set => duongDanAnh = value; }
        public string Ma_MonHoc1 { get => Ma_MonHoc; set => Ma_MonHoc = value; }
        public string DiaChi1 { get => DiaChi; set => DiaChi = value; }
    }
}
