using System;
using System.Collections.Generic;
using System.Text;
using System.Data.OleDb;
using System.Windows.Forms;

namespace Import2DataBaseFormExcel.SourceCode
{
    class Connection
    {
        private static System.Data.OleDb.OleDbConnection con;
        public static void conect()
        {
            con = new OleDbConnection(@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\qkhanh\excell");
        }
        public static System.Data.DataTable get_struct(string lenh)
        {
            System.Data.OleDb.OleDbDataAdapter thichung = new OleDbDataAdapter(lenh, con);
            System.Data.DataTable bang = new System.Data.DataTable();
            thichung.Fill(bang);
            return bang;
        }
        public static Boolean Exe_SQL(String lenh)
        {
            Boolean dung = new Boolean();
            dung = true;
            System.Data.OleDb.OleDbCommand xl_lenh = new OleDbCommand();
            try
            {
                con.Open();
                xl_lenh.Connection = con;
                xl_lenh.CommandText = lenh;
                int i = xl_lenh.ExecuteNonQuery();
                con.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi: " + ex.ToString() + lenh + " : ");
                dung = false;
                con.Close();
            }
            return dung;
        }
    }
}
