using System;
using System.Collections.Generic;
using System.IO.Ports;
using System.Text;
using System.Drawing;
using System.Windows.Forms;
using System.Threading;

namespace WindowsFormsApplication1
{
    public class clsScanner
    {
        public event SerialDataReceivedEventHandler Datareceived;
        SerialPort Scanner;
        Form1 _frm;
        //clsdataconvert dataconvert;

        private string _data;

        public string Data
        {
            get { return _data; }
            set { _data = value; }
        }

        private string _COMnum;

        public string COMnum
        {
            get { return _COMnum; }
            set { _COMnum = value; }
        }

        public clsScanner(Form1 frm)
        {
            _frm = frm;
            Scanner = new SerialPort();
            //dataconvert = new clsdataconvert();
        }

        public bool ketnoi(ToolStripStatusLabel lb)
        {
            try
            {
                Scanner.PortName = _COMnum;
                Scanner.BaudRate = 9600;
                Scanner.DataBits = 8;
                Scanner.ReadBufferSize = 1024;
                Scanner.WriteBufferSize = 512;
                Scanner.Parity = Parity.None;
                Scanner.StopBits = StopBits.One;
                Scanner.Handshake = Handshake.None;
                //Scanner.DtrEnable = true;
                Scanner.DataReceived += Scanner_DataReceived;
                Scanner.Open();
                lb.BackColor = Color.Green;
                return true;
            }
            catch (Exception)
            {
                Scanner.Close();
                lb.BackColor = Color.Red;
                return false;
            }

        }

        public void ngatketnoi()
        {
            try
            {
                Scanner.Close();
            }
            catch (Exception)
            {

            }
        }


        public void ReadData()
        {
            Scanner.WriteLine("LON\r");
            Thread.Sleep(300);
            //Scanner.WriteLine("LOFF\r\n");
        }

        public TextBox tb;
        private void Scanner_DataReceived(object sender, SerialDataReceivedEventArgs e)
        {
            try
            {

                _data = "";
                _data = Scanner.ReadLine();

                if (Datareceived != null)
                {
                    _frm.tb_barcode.Text = _data;
                    Datareceived(this, e);
                }
            }
            catch (Exception)
            {
                //throw;
            }
        }


    }
}
