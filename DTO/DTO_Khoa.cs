using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DTO
{
    public class DTO_Khoa
    {
        private string ma_Khoa;
        private string TenKhoa;

        public DTO_Khoa()
        {
        }

        public DTO_Khoa(string ma_Khoa, string tenKhoa)
        {
            this.ma_Khoa = ma_Khoa;
            TenKhoa = tenKhoa;
        }

        public string Ma_Khoa { get => ma_Khoa; set => ma_Khoa = value; }
        public string TenKhoa1 { get => TenKhoa; set => TenKhoa = value; }
    }
}
