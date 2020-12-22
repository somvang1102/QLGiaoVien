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
namespace WindowsFormsApp1
{
    public partial class thongtintaikhoan : Form
    {
        Bus_DangNhap b = new Bus_DangNhap();
        private string tenTaiKhoan;
        DataTable dt;
        public string TenTaiKhoan { get => tenTaiKhoan; set => tenTaiKhoan = value; }
        public thongtintaikhoan()
        {
            InitializeComponent();
        }

        private void thongtintaikhoan_Load(object sender, EventArgs e)
        {
            txtTentk.Enabled = false;
            txtTenND.Enabled = false;
            txtEmail.Enabled = false;
            txtType.Enabled = false;
            txtTentk.Text = tenTaiKhoan;
            dt = b.getAccount(tenTaiKhoan);
            txtEmail.Text = dt.Rows[0]["Email"].ToString();
            txtTenND.Text = dt.Rows[0]["DisplayName"].ToString();
            int x = Convert.ToInt32(dt.Rows[0]["Type"].ToString());
            if(x==1)
            {
                txtType.Text = "Quản trị viên";
                
            }
            else
            {
                txtType.Text = "Giáo viên";
            }
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox4_TextChanged(object sender, EventArgs e)
        {

        }

        private void btnok_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
