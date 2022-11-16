using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TranMinhDuc_201210096
{
    internal class DBConfig
    {
        private string STRING_CONNECT = "";
        private SqlDataAdapter sqlDataAdapter; // Dùng để chèn dữ liệu vào DATA Table hoặc DataSet
        private SqlCommand sqlCommand;// Thực thi câu lệnh truy vấn

        public DataTable table(string query)
        {
            DataTable dataTable = new DataTable();
            using (SqlConnection sqlConnection = new SqlConnection(STRING_CONNECT))
            {
                sqlConnection.Open();
                sqlDataAdapter = new SqlDataAdapter(query, sqlConnection);
                sqlDataAdapter.Fill(dataTable);
                sqlConnection.Close();
            }
            return dataTable;

        }

        public void Excute(string query) // Hàm dùng để update ,insert và delete
        {
            using (SqlConnection sqlConnection = new SqlConnection(STRING_CONNECT))
            {
                sqlConnection.Open();
                sqlCommand = new SqlCommand(query, sqlConnection);
                sqlCommand.ExecuteNonQuery(); // thực thi câu truy vấn
                sqlConnection.Close();
            }
        }

        public object GetValue(string query)
        {
            object val;
            using (SqlConnection sqlConnection = new SqlConnection(STRING_CONNECT))
            {
                sqlConnection.Open();
                sqlCommand = new SqlCommand(query, sqlConnection);
                val = sqlCommand.ExecuteScalar();
                sqlConnection.Close();
            }
            return val;
        }
    }
}
