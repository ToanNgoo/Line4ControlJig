using System;
using System.Collections.Generic;
using System.Net;
using System.Net.Sockets;
using System.IO.Ports;
using System.Text;
using System.Threading;
using System.Windows.Forms;
using System.Drawing;

namespace WindowsFormsApplication1
{
    public class clsCOM_MES
    {
        SerialPort DVM;
        clsdataconvert dataconvert;
        //clsdataconvert dataconvert;
        private string _COMnum;
        
        public string COMnum
        {
            get { return _COMnum; }
            set { _COMnum = value; }
        }

        private bool _flag_smart;
        public bool flag_smart
        {
            get { return _flag_smart; }
            set { _flag_smart = value; }
        }

        private bool _flag_pmp;
        public bool flag_pmp
        {
            get { return _flag_pmp; }
            set { _flag_pmp = value; }
        }

        private bool _flag_misum;
        public bool flag_misum
        {
            get { return _flag_misum; }
            set { _flag_misum = value; }
        }

        private bool _flag_5cell;
        public bool flag_5cell
        {
            get { return _flag_5cell; }
            set { _flag_5cell = value; }
        }

         private string _datasent;
        public string datasent
        {
            get { return _datasent; }
            set { _datasent = value; }
        }

        private string _data;
        public string Data
        {
            get { return _data; }
            set { _data = value; }
        }

        public clsCOM_MES()
        {
            DVM = new SerialPort();
            dataconvert = new clsdataconvert();
            //dataconvert = new clsdataconvert();
           
        }

        public void Disconnect()
        {
            try
            {
                DVM.Close();
            }
            catch (Exception)
            {
                ;
            }
        }

        public bool ketnoi(string baurate)
        {
            try
            {
                DVM.PortName = _COMnum;
                DVM.BaudRate = int.Parse(baurate);
                DVM.DataBits = 8;
                DVM.ReadBufferSize = 1024;
                DVM.WriteBufferSize = 512;
                DVM.Handshake = Handshake.None;
                DVM.Parity = Parity.None;
                DVM.DtrEnable = true;
                DVM.DataReceived += DVM_DataReceived;
                DVM.Open();
                return true;
            }
            catch (Exception)
            {
                DVM.Close();
                return false;
            }

        }



        public string UpJig = "";

        void DVM_DataReceived(object sender, SerialDataReceivedEventArgs e)
        {          
            _data = DVM.ReadExisting();
            if ((_data == "END!") || (_data == "EROR"))
            {
                UpJig = "UP";
            }

            if(_flag_smart)
            {
                //_data = DVM.ReadExisting();
                if (_data == "IDN?")
                {
                    DVM.Write("???");

                }
            }
            else if (_flag_pmp)
            {
                //_data = DVM.ReadExisting();
                if (_data == "IDN?")
                {
                    DVM.Write("V100");
                }
                
            }
            else if (_flag_misum)
            {
                byte[] datrev = new byte[100];
                datrev = System.Text.Encoding.GetEncoding(1252).GetBytes(DVM.ReadExisting());

                string a = "";
                for(int i=0;i<datrev.Length;i++)
                {
                    a = a + dataconvert.hex_str2(datrev[i]);
                }
                _data = a;
                _data = dataconvert.hex2ascii(a);
            }
            else if (_flag_5cell)
            {
                //_data = DVM.ReadExisting();
                if (_data == "IDN?")
                {
                    DVM.Write("V120");

                }
            }

            
        }
        public void sent_data(string data)
        {
            DVM.Write(data);
        }
        public void sent_data_byte(string factor1,string factor2,string factor3,string factor4)
        {
            byte[] Txsent;
            Txsent = new byte[50];
            Txsent[0] = 0x02;
            Txsent[1] = byte.Parse(dataconvert.H2D(factor1).ToString());
            Txsent[2] = byte.Parse(dataconvert.H2D(factor2).ToString());
            Txsent[3] = byte.Parse(dataconvert.H2D(factor3).ToString());
            Txsent[4] = byte.Parse(dataconvert.H2D(factor4).ToString());
            Txsent[5] = 0x03;
            DVM.Write(Txsent, 0, 50);

        }

        public void ngatketnoi()
        {
            try
            {
                DVM.Close();
            }
            catch (Exception)
            {

            }
        }

       
    }
}
