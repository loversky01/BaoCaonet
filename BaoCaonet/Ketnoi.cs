using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BaoCaonet
{
    internal class Ketnoi
    {
        public static string ConnectString = @"Data Source=LOVERSKY\SQLEXPRESS;Initial Catalog=QuanLyChungCu;Integrated Security=True";
        public static DataTable SelectDB(string sql)
        {
            using (SqlConnection conn = new SqlConnection(ConnectString))
            {
                using (SqlDataAdapter dad = new SqlDataAdapter(sql, conn))
                {
                    using (DataSet dts = new DataSet())
                    {
                        dad.Fill(dts);
                        return dts.Tables[0];
                    }
                }
            }

           
        }

        public static void UpInDeDB(string sql)
        {
            using (SqlConnection conn = new SqlConnection(ConnectString))
            {
                if (conn.State == ConnectionState.Open)
                {
                    conn.Close();
                }
                conn.Open();
                using (SqlCommand cmd = new SqlCommand(sql, conn))
                {
                    cmd.CommandType = CommandType.Text;
                    cmd.ExecuteNonQuery();
                    cmd.Dispose();
                }
                conn.Close();
                conn.Dispose();
            }
        }
    }
}
