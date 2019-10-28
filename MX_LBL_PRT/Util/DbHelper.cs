using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MX_LBL_PRT.Util
{
    public class DbHelper
    {
        public DbHelper()
        {
            
        }

        public SqlConnection GetSqlConnection()
        {
            return new SqlConnection(@"Data Source=192.168.16.122\MXSQL01;Initial Catalog=MXSQLDB01;Persist Security Info=True;User ID=MX_LBL_PRT;Password=Bizlink@2019");
        }
    }
}
