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
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;
namespace WindowsFormsApp1
{
    public partial class FormQuanLy : Form
    {
        
        public FormQuanLy()
        {
            
        }
        //void loadDtgvGV()
        //{
        //    dgvgiaovien.DataSource =giaoVien.getGiaoVien();
        //    dgvgiaovien.Columns[0].HeaderText = "Mã giáo viên";
        //    dgvgiaovien.Columns[1].HeaderText = "Họ tên";
        //    dgvgiaovien.Columns[2].HeaderText = "Ngày sinh";
        //    dgvgiaovien.Columns[3].HeaderText = "Giới tính";
        //    dgvgiaovien.Columns[4].HeaderText = "Email";
        //    dgvgiaovien.Columns[5].HeaderText = "Học vị";
        //    dgvgiaovien.Columns[6].HeaderText = "Học hàm";
        //    dgvgiaovien.Columns[7].HeaderText = "Tên môn giảng dạy";
        //    dgvgiaovien.Columns[8].HeaderText = "Chức vụ";
        //    dgvgiaovien.Columns[9].HeaderText = "Bộ môn";
        //    dgvgiaovien.Columns[10].HeaderText = "Khoa";
            

        //}
        void loadDtgvGV2()
        {
            


        }
        void loadMa_Khoa()
        {
            
        }
        void loadHocHam()
        {
            
        }
        void loadHocVi()
        {
           
        }

        public List<DTO_GiaoVien> loadGiaoVien()
        {
           
        }
        bool checkMaGV()
        {
            
        }
        void loadBoMon(string s)
        {
            
        }
        void getMonHoc(string s)
        {
            
        }
        private void FormQuanLy_Load(object sender, EventArgs e)
        {
            
        }

        private void groupBox1_Enter(object sender, EventArgs e)
        {

        }

        private void btnFresh_Click(object sender, EventArgs e)
        {
           
        }

        private void dgvgiaovien_CellClick(object sender, DataGridViewCellEventArgs e)
        {

            
        }
   
        private void cmbkhoa_SelectedIndexChanged(object sender, EventArgs e)
        {
            loadBoMon(this.cmbkhoa.Text);
        }

        private void btnThem_Click(object sender, EventArgs e)
        {
           
        }
       
        //private void btnChonAnh_Click(object sender, EventArgs e)
        //{

        //    if (op.ShowDialog() == DialogResult.OK)
        //    {
        //        s = op.FileName;
        //        pbanhDaiDien.Image = Image.FromFile(s);
        //    }
           
               
            
        //}
        
        //private void btnThayAnh_Click(object sender, EventArgs e)
        //{
        //    if(giaoVien.ThemAnh(s,txtmagv.Text)==true)
        //    {
        //        MessageBox.Show(" thành công");
        //    }
        //    else
        //    {
        //        MessageBox.Show("thay đổi thất bại");
        //    }
        //}

        private void cmbbomon_SelectedIndexChanged(object sender, EventArgs e)
        {
            
        }

        private void btnXoa_Click(object sender, EventArgs e)
        {
           
            
        }

        private void button2_Click(object sender, EventArgs e)
        {
           
        }
        //hàm in ra excel
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
        private void btnIn_Click(object sender, EventArgs e)
        {
            ExportFile("Danh Sách giáo viên", dgvgiaovien);
        }

        private void groupBox2_Enter(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            
        }

        private void txtmagv_KeyPress(object sender, KeyPressEventArgs e)
        {
           
        }
    }
}
