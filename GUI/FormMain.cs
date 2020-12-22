using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using DTO;
using QLGV;
using Excel = Microsoft.Office.Interop.Excel;
namespace WindowsFormsApp1
{
    public partial class FormMain : Form
    {
        

        public FormMain()
        {
            InitializeComponent();
        }

        private void FormMain_Load(object sender, EventArgs e)
        {
            comboBoxKhoa.Visible = false;
            this.Controls.Add(data1);
            textBox1.Text = tenTaiKhoan;

            data1.Visible = false;
            //comboBoxKhoa.Visible = false;
            button1.Visible = false;
            btnBoMon.Visible = false;
        }
        
        private void khoaToolStripMenuItem_Click(object sender, EventArgs e)
        {
            
            
        }

       

        private void bộMônToolStripMenuItem_Click(object sender, EventArgs e)
        {
            
        }

        private void comboBoxKhoa_SelectedIndexChanged(object sender, EventArgs e)
        {
            
            
        }
        // hàm xuất file excel 
        public static void ExportFile(string Header, DataGridView dgv)
        {
            // Tạo đối tượng mở Explorer
            SaveFileDialog fsave = new SaveFileDialog();
            // Chỉ ra đuôi của tệp tin
            fsave.Filter = "(Tất cả các tệp)|*.*|(Các tệp excel)|*.xlsx";
            fsave.ShowDialog();

            if (fsave.FileName != "")
            {
                // Tạo Excel App
                Excel.Application app = new Excel.Application();
                // Tạo 1 workbook
                Excel.Workbook wb = app.Workbooks.Add(Type.Missing);
                Excel.Worksheet sheet = null;
                try
                {
                    // Đọc dữ liệu
                    sheet = wb.ActiveSheet;
                    sheet.Range[sheet.Cells[1, 1], sheet.Cells[1, dgv.ColumnCount]].Merge();
                    sheet.Cells[1, 1].Value = Header;
                    sheet.Cells[1, 1].Font.Name = "Times New Roman";
                    sheet.Cells[1, 1].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                    sheet.Cells[1, 1].Font.Size = 20;
                    sheet.Cells[1, 1].Borders.Weight = Excel.XlBorderWeight.xlThin;
                    //Sinh tiêu đề
                    for (int i = 1, k = 1; i <= dgv.Columns.Count; i++)
                    {
                        if (dgv.Columns[i - 1].Visible == false) continue;
                        sheet.Cells[2, k] = dgv.Columns[i - 1].HeaderText;
                        sheet.Cells[2, k].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                        sheet.Cells[2, k].Font.Name = "Times New Roman";
                        sheet.Cells[2, k].Font.Bold = true;
                        sheet.Cells[2, k].Borders.Weight = Excel.XlBorderWeight.xlThin;
                        k++;
                    }
                    // Sinh dữ liệu
                    for (int i = 1; i <= dgv.RowCount - 1; i++)
                    {
                        if (dgv.Columns[0].Visible == false) continue;
                        sheet.Cells[i + 2, 1] = dgv.Rows[i - 1].Cells[0].Value;
                        sheet.Cells[i + 2, 1].Font.Name = "Times New Roman";
                        sheet.Cells[i + 2, 1].Borders.Weight = Excel.XlBorderWeight.xlThin;
                        for (int j = 2, k = 2; j <= dgv.Columns.Count; j++)
                        {
                            if (dgv.Columns[j - 1].Visible == false) continue;
                            sheet.Cells[i + 2, k] = dgv.Rows[i - 1].Cells[j - 1].Value;
                            sheet.Cells[i + 2, k].Font.Name = "Times New Roman";
                            sheet.Cells[i + 2, k].Borders.Weight = Excel.XlBorderWeight.xlThin;
                            k++;
                        }
                    }
                    sheet.Columns.AutoFit();
                    wb.SaveAs(fsave.FileName);
                    MessageBox.Show("Ghi thành công!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                finally
                {
                    app.Quit();
                    wb = null;
                }

            }
            else
            {
                MessageBox.Show("Bạn không chọn tệp tin nào", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
          ExportFile("DANH SÁCH BÀI BÁO KHOA HỌC",data1 );
        }

        private void đăngXuấtToolStripMenuItem_Click(object sender, EventArgs e)
        {
            
            DangNhap f = new DangNhap();
            this.Hide();
            f.ShowDialog();
            this.Close();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            data1.DataSource = bus_Khoa.getAllBoMon();
        }

        private void đổiMậtKhẩuToolStripMenuItem_Click(object sender, EventArgs e)
        {
           
            
            


        }

        private void họcHàmToolStripMenuItem_Click(object sender, EventArgs e)
        {
            data1.Visible = true;
            button1.Visible = true;
            btnBoMon.Visible = false;
            comboBoxKhoa.Visible = false;
            data1.DataSource = bus_HocHam_HocVi.thongkeHocHam("SELECT dbo.GiaoVien.HoTen,dbo.HocHam.Ma_HocHam,tenHocHam FROM dbo.HocHam,dbo.GiaoVien WHERE dbo.GiaoVien.Ma_HocHam = dbo.HocHam.Ma_HocHam");
            data1.Columns[0].HeaderText = "Giáo Viên";
            data1.Columns[1].HeaderText = "Mã Học Hàm";
            data1.Columns[2].HeaderText = "Tên Học Hàm";
        }

        private void họcVịToolStripMenuItem_Click(object sender, EventArgs e)
        {
            data1.Visible = true;
            button1.Visible = true;
            data1.DataSource = bus_HocHam_HocVi.thongkeHocHam("SELECT dbo.GiaoVien.HoTen,dbo.HocVi.Ma_HocVi,TenHocVi FROM dbo.HocVi,dbo.GiaoVien WHERE dbo.GiaoVien.Ma_HocVi = dbo.HocVi.Ma_HocVi");
            data1.Columns[0].HeaderText = "Giáo Viên";
            data1.Columns[1].HeaderText = "Mã Học Vị";
            data1.Columns[2].HeaderText = "Tên Học Vị";
        }

        private void data1_CellClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void danhSáchGiáoViênToolStripMenuItem_Click(object sender, EventArgs e)
        {
            FormQuanLy f = new FormQuanLy();
             dt = bus_DangNhap.getAccount(textBox1.Text);
            f.LoaiTaiKhoan = Convert.ToInt32(dt.Rows[0]["Type"].ToString());
            f.ShowDialog();
        }

        private void thôngTinToolStripMenuItem_Click(object sender, EventArgs e)
        {
            thongtintaikhoan f = new thongtintaikhoan();
            f.TenTaiKhoan = textBox1.Text;
           
            f.ShowDialog();
   

        }

        private void quảnLýTàiKhoảnToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void lịchGiảngDạyToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ThongKeGiaoVien f = new ThongKeGiaoVien();
            f.ShowDialog();


        }
    }
}
