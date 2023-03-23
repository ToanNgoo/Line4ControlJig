using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Threading;
using System.Windows.Forms;
using System.IO;

namespace WindowsFormsApplication1
{
    public partial class Form2 : Form
    {
        int dem = 0;
        clsLocaldb localdb;
        clsLocaldb localdb1;
        string link, link1;
        Form1 frm;
        public Form2(string _link, string _link1, Form1 _frm)
        {
            InitializeComponent();
            link = _link;
            link1 = _link1;
            frm = _frm;
        }    
        
       
        private void Form2_Load(object sender, EventArgs e)
        {                      
            localdb = new clsLocaldb(link);
            localdb1 = new clsLocaldb(link1);       
            groupBox13.Text = frm.code_Jig;
        }      

        private void lb_Jig_TextChanged(object sender, EventArgs e)
        {
            progressBar1.Value = dem;
            if (dem < 30000)
                progressBar1.ForeColor = Color.Green;
            else if (dem >= 30000 && dem < 40000)
                progressBar1.ForeColor = Color.Orange;
            else
                progressBar1.ForeColor = Color.Red;
        }

        bool check = true;
        private void timer1_Tick(object sender, EventArgs e)
        {
            if (!frm.frm2 && check)
            {
                check = false;
                frm.frm2 = true;
                if (frm._checkServer)
                {
                    lb_Jig.Text = frm._dem;
                    groupBox13.Text = frm.code_Jig + "/Server";
                }
                else
                {
                    lb_Jig.Text = frm._demLocal;
                    groupBox13.Text = frm.code_Jig + "/Local";
                }
                dem = int.Parse(lb_Jig.Text);
                check = true;
            }
        }
    }
}
