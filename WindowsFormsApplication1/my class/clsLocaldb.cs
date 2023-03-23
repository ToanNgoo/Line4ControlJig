using System;
using System.Collections.Generic;
using System.Text;
using System.Windows.Forms;
using System.Data;
using System.Data.OleDb;
using System.Threading;
namespace WindowsFormsApplication1
{
    class clsLocaldb
    {    
        string constr = "";       
        bool _flag_recei_line;
        public bool flag_recei_line
        {
            get { return _flag_recei_line; }
            set { _flag_recei_line = value; }
        }

        public clsLocaldb(string link)
        {            
            constr = @"Provider=Microsoft.Jet.OLEDB.4.0; Data Source = " + link + @"\Database.mdb";
        }

        public DataTable getData(string str)
        {
            DataTable dt = new DataTable();
            OleDbDataAdapter da = new OleDbDataAdapter(str, constr);
            da.Fill(dt);
            return dt;
        }
        
        public void uploadCounting(string model, int Num_Jig, int count)
        {
            OleDbConnection cnn = new OleDbConnection(constr);
            cnn.Open();
            string str = " UPDATE Data SET Data.Count = '" + count.ToString() + "' where Data.model = '" + model + "' and Data.Jig = '" + Num_Jig.ToString() + "'";
            OleDbCommand cmd = new OleDbCommand(str, cnn);
            cmd.ExecuteNonQuery();
            cnn.Close();
        }

        public void uploadCounting(string codeJig, int count)
        {
            OleDbConnection cnn = new OleDbConnection(constr);
            cnn.Open();
            string str = "UPDATE Data SET Data.Count = '" + count.ToString() + "' where Data.Code_Jig = '" + codeJig + "'";
            OleDbCommand cmd = new OleDbCommand(str, cnn);
            cmd.ExecuteNonQuery();
            cnn.Close();
        }

        public int loadCounting(string model, int Num_Jig)
        {
            string str = "select Count from Data where model = '" + model + "' and Jig = '" + Num_Jig + "'";
            DataTable dt = new DataTable();
            dt = getData(str);
            int count = 0;
            foreach (DataRow dr in dt.Rows)
            {
                count = int.Parse(dr.ItemArray[0].ToString());
            }
            return count;
        }


        public string[] loadProgramFCT(string barcode)
        {
            string str = "Select Distinct Program_FCT from Data";
            DataTable dt = new DataTable();
            dt = getData(str);
            string[] data = new string[dt.Rows.Count];
            int count = 0;
            foreach (DataRow dr in dt.Rows)
            {
                if(dr.ItemArray[0].ToString() != "")
                {
                    data[count] = dr.ItemArray[0].ToString();
                    count++;
                }               
            }
            return data;
        }


        public int loadCounting(string codeJig)
        {
            string str = "select Count from Data where Code_Jig = '" + codeJig + "'";
            DataTable dt = new DataTable();
            dt = getData(str);
            int count = 0;
            foreach (DataRow dr in dt.Rows)
            {
                count = int.Parse(dr.ItemArray[0].ToString());
            }
            return count;
        }

        public string[] getModel()
        {
            string str = "select Model from Data;";
            DataTable dt = new DataTable();
            dt = getData(str);
            string[] model = new string[dt.Rows.Count];
            int count = 0;
            foreach (DataRow dr in dt.Rows)
            {
                model[count] = dr.ItemArray[0].ToString();
                count++;
            }
            return model;
        }




        public string _getModel(string code_Jig)
        {
            try
            {
                string str = "select Model from Data where Code_jig = '" + code_Jig + "';";
                DataTable dt = new DataTable();
                dt = getData(str);
                string[] model = new string[dt.Rows.Count];
                int count = 0;
                foreach (DataRow dr in dt.Rows)
                {
                    model[count] = dr.ItemArray[0].ToString();
                }
                return model[0];
            }
            catch (Exception )
            {
                return "";
            }
            
        }

        public string _getProgram(string code_Jig)
        {
            try
            {
                string str = "select Program_FCT from Data where Code_jig = '" + code_Jig + "';";
                DataTable dt = new DataTable();
                dt = getData(str);
                string[] model = new string[dt.Rows.Count];
                int count = 0;
                foreach (DataRow dr in dt.Rows)
                {
                    model[count] = dr.ItemArray[0].ToString();
                }
                return model[0];
            }
            catch (Exception)
            {
                return "";
            }
            
        }

        public bool Status_QR(string code_Jig)
        {
            try
            {
                string str = "select QR_Code from Data where Code_jig = '" + code_Jig + "';";
                DataTable dt = new DataTable();
                dt = getData(str);
                string[] model = new string[dt.Rows.Count];
                foreach (DataRow dr in dt.Rows)
                {
                    model[0] = dr.ItemArray[0].ToString();
                }
                if (model[0] == "Yes")
                {
                    return true;
                }
                else
                {
                    return false;
                }
            }
            catch (Exception)
            {
                return false;
            }

        }

        public void uploadmodel(string code_jig, string FCT_program, string status)
        {
            OleDbConnection cnn = new OleDbConnection(constr);
            cnn.Open();
            string str = "INSERT INTO Data VALUES ('" + code_jig + "','" + FCT_program + "','" + "0" + "','" + status + "')";
            OleDbCommand cmd = new OleDbCommand(str, cnn);
            cmd.ExecuteNonQuery();
            cnn.Close();
        }

    }
}