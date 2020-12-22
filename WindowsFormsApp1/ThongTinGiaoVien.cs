using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using QLGV;
using DTO;
using System.Data;
using System.Data.SqlClient;
namespace WindowsFormsApp1
{
    public partial class ThongTinGiaoVien : Form
    {
        Bus_GiaoVien giaovien2 = new Bus_GiaoVien();
        private string maGiaVien;
        private string duongDanAnh;


        public ThongTinGiaoVien()
        {
            InitializeComponent();
        }

        public string MaGiaVien { get => maGiaVien; set => maGiaVien = value; }
       // public string DuongDanAnh { get => duongDanAnh; set => duongDanAnh = value; }

        private void ThongTinGiaoVien_Load(object sender, EventArgs e)
        {
           

        }

        private void txtBoMon_TextChanged(object sender, EventArgs e)
        {

        }
    }
}
