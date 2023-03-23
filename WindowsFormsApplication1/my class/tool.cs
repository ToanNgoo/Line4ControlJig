using System;
using System.Collections.Generic;
using System.Text;
using System.Windows.Forms;
using System.Drawing;

namespace Import2DataBaseFormExcel.SourceCode
{
    class tool
    {
        public static void style_dgview(System.Windows.Forms.DataGridView dataGridView1)
        {
            //if (dem > 48000)
            //{
            //    if (dem > 48000 && dem < 49000)
            //    {
            //        if (int.Parse(DateTime.Now.Second.ToString()) == 1 && int.Parse(DateTime.Now.Minute.ToString()) == 0 && trangThai == false)
            //        {
            //            PLC.Writeplc("M113", 1);
            //            DialogResult result = new System.Windows.Forms.DialogResult();
            //            result = MessageBox.Show("S?p ©¢?n chu k? thay kim h?y b?o ME, PE", "question", MessageBoxButtons.YesNo);
            //            if (result == DialogResult.Yes)
            //            {
            //                PLC.Writeplc("M113", 0);
            //            }
            //            else
            //            {

            //            }
            //        }
            //    }
            //    else if (dem > 49000 && dem < 50000)
            //    {
            //        if (int.Parse(DateTime.Now.Second.ToString()) % 30 == 0 && int.Parse(DateTime.Now.Minute.ToString()) == 0 && trangThai == false)
            //        {
            //            PLC.Writeplc("M113", 1);
            //            DialogResult result = new System.Windows.Forms.DialogResult();
            //            result = MessageBox.Show("S?p ©¢?n chu k? thay kim h?y b?o ME, PE", "question", MessageBoxButtons.YesNo);
            //            if (result == DialogResult.Yes)
            //            {
            //                PLC.Writeplc("M113", 0);
            //            }
            //            else
            //            {

            //            }
            //        }
            //    }
            //    else if (dem > 50000)
            //    {
            //        MessageBox.Show("?? h?t h?n s? d?ng kim");
            //        kq1 = false;
            //        PLC.Writeplc("M108", 0);
            //        PLC.Writeplc("M111", 1);
            //    }

            //}

            DataGridViewCellStyle style1 = new DataGridViewCellStyle();
            style1.ForeColor = Color.Blue;
            style1.BackColor = Color.Linen;
            DataGridViewCellStyle style2 = new DataGridViewCellStyle();
            style2.ForeColor = Color.Red;
            style2.BackColor = Color.White;
            for (int i = dataGridView1.RowCount - 1; i >= 0; i--)
            {
                if (i % 2 == 0)
                {
                    dataGridView1.Rows[i].DefaultCellStyle = style1;
                }
                else if (i % 2 != 0)
                {
                    dataGridView1.Rows[i].DefaultCellStyle = style2;

                }
            }
        }
    }
}
