using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.SqlClient;
using System.Data;
namespace DAL
{
    public class DAL_con
    {
        
            protected SqlConnection conn = new SqlConnection(@"Data Source=.\SQLEXPRESS;Initial Catalog=QuanLyGiaoVien;Integrated Security=True");
        public  bool executeQuery(string query )
        {
            try
            {
                SqlCommand command = new SqlCommand(query, conn);
                command.ExecuteNonQuery();
                return true;
            }
            catch (Exception ex)
            {
                return false;
            }
        }
        
        public  DataTable GetDataTable(string query)
        {
            
            DataTable data = new DataTable();
            SqlDataAdapter adapter = new SqlDataAdapter(query, conn);
            adapter.Fill(data);
            return data;
        }
    }

     
}
