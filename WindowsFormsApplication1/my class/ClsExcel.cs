using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel;
using System.Runtime.InteropServices;
using System.Data.OleDb;
using System.Data;
using System.Windows.Forms;
using System.IO;

//using Excel = Microsoft.Office.Interop.Excel;
namespace WindowsFormsApplication1
{
    class ClsExcel
    {       
        public DataTable ReadLog(string path)
        {
            DataTable dtb = new DataTable();
            int tmp = 0;
            string Fulltext;

            _let:
            try
            {
                using (StreamReader sr = new StreamReader(path))
                {
                    while (!sr.EndOfStream)
                    {
                        Fulltext = sr.ReadToEnd().ToString();
                        string[] rows = Fulltext.Split('\r');

                        for (int i = 6; i < rows.Length; i++)
                        {
                            string[] rowValues = rows[i].Split(',');
                            if (i == 6)
                            {
                                for (int j = 0; j < rowValues.Count(); j++)
                                {
                                    dtb.Columns.Add(rowValues[j]);
                                }
                            }
                            else
                            {
                                DataRow dr = dtb.NewRow();
                                for (int k = 0; k < rowValues.Count(); k++)
                                {
                                    dr[k] = rowValues[k].ToString();
                                }
                                dtb.Rows.Add(dr);
                            }
                        }
                    }
                    tmp = 0;
                    sr.Close();
                }
                
            }
            catch (Exception)
            {
                tmp = 1;
            }

            if (tmp == 1)
                goto _let;
            
            return dtb;
        }


        public DataTable ReadCSVFile(string path, int RowStart)
        {
            DataTable dtb = new DataTable();
            int tmp = 0;
            string Fulltext;

        _let:
            try
            {
                using (StreamReader sr = new StreamReader(path))
                {
                    while (!sr.EndOfStream)
                    {
                        Fulltext = sr.ReadToEnd().ToString();
                        string[] rows = Fulltext.Split('\r');

                        for (int i = RowStart; i < rows.Length; i++)
                        {
                            string[] rowValues = rows[i].Split(',');
                            if (i == RowStart)
                            {
                                for (int j = 0; j < rowValues.Count(); j++)
                                {
                                    dtb.Columns.Add(rowValues[j]);
                                }
                            }
                            else
                            {
                                DataRow dr = dtb.NewRow();
                                for (int k = 0; k < rowValues.Count(); k++)
                                {
                                    dr[k] = rowValues[k].ToString();
                                }
                                dtb.Rows.Add(dr);
                            }
                        }
                    }
                    tmp = 0;
                }
            }
            catch (Exception)
            {
                tmp = 1;
            }

            if (tmp == 1)
                goto _let;

            return dtb;
        }

        }
    }

