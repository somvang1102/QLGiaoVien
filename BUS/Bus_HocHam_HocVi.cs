using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DAL;
using System.Data;
using System.Data.SqlClient;
using DTO;
namespace QLGV
{
    
    public class Bus_HocHam_HocVi
    {
        DAL_HocHam dAL_HocHam = new DAL_HocHam();
        public DataTable  thongkeHocHam(string s)
        {
            return dAL_HocHam.thongkeHocHam(s);
        }
    }
}
