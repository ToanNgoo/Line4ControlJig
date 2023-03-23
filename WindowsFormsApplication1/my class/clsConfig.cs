using System;
using System.Collections.Generic;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.IO.Ports;
using System.Threading;

namespace WindowsFormsApplication1
{
    public class clsConfig
    {
        public void loadlistcom(ComboBox cb)
        {
            cb.Items.Clear();
            string[] ports = SerialPort.GetPortNames();
            foreach (string port in ports)
            {
                cb.Items.Add(port);
            }
        }      

        public void loadconfig_setup(TextBox IP, TextBox Link_1, TextBox Link_2, TextBox Link_3, TextBox Link_4,TextBox Link_5, TextBox Link_6, TextBox Ethernet, ComboBox Scaner_Jig)
        {
            string[] data = null;
            string str;
            FileStream FS = new FileStream(Application.StartupPath + @"\Setup.ini", FileMode.Open);
            StreamReader SR = new StreamReader(FS);
            while (SR.EndOfStream == false)
            {
                str = SR.ReadLine();
                data = str.Split('=');

                switch (data[0])
                {
                    case "IP server":
                        IP.Text = data[1];
                        break;    
                    case"Link 1":
                        Link_1.Text = data[1];
                        break;
                    case "Link 2":
                        Link_2.Text = data[1];
                        break;
                    case "Link 3":
                        Link_3.Text = data[1];
                        break;
                    case "Link 4":
                        Link_4.Text = data[1];
                        break;
                    case "Link 5":
                        Link_5.Text = data[1];
                        break;
                    case "Link 6":
                        Link_6.Text = data[1];
                        break;
                    case "Ethernet":
                        Ethernet.Text = data[1];
                        break;
                    case "Scaner Jig":
                        Scaner_Jig.Text = data[1];
                        break;
                    default:
                        break;
                }
            }
            SR.Close();
            FS.Close();
        }

        public void loaduser(string user, string pass)
        {
            string[] data = null;
            string str;
            FileStream FS = new FileStream(Application.StartupPath + @"\Config\User.txt", FileMode.Open);
            StreamReader SR = new StreamReader(FS);
            while (SR.EndOfStream == false)
            {
                str = SR.ReadLine();
                data = str.Split('=');

                switch (data[0])
                {
                    case "User":
                        user = data[1];
                        break;
                    case "Pass":
                        pass = data[1];
                        break;                   
                }
            }
            SR.Close();
            FS.Close();
        }

        public void saveconfig_setup(TextBox IP, TextBox Link_1, TextBox Link_2, TextBox Link_3, TextBox Link_4,TextBox Link_5, TextBox Link_6, TextBox Ethernet, ComboBox Scaner_Jig)
        {
            FileStream FS = new FileStream(Application.StartupPath + @"\Setup.ini", FileMode.Create);
            StreamWriter SW = new StreamWriter(FS);

            SW.WriteLine("IP server=" + IP.Text);
            SW.WriteLine("Link 1=" + Link_1.Text);
            SW.WriteLine("Link 2=" + Link_2.Text);
            SW.WriteLine("Link 3=" + Link_3.Text);
            SW.WriteLine("Link 4=" + Link_4.Text);
            SW.WriteLine("Link 5=" + Link_5.Text);
            SW.WriteLine("Link 6=" + Link_6.Text);
            SW.WriteLine("Ethernet=" + Ethernet.Text);
            SW.WriteLine("Scaner Jig=" + Scaner_Jig.Text);
            SW.Close();
            FS.Close();
        }

        public void saveconfig_device(ComboBox scan1)
        {
            FileStream FS = new FileStream(Application.StartupPath + @"\Device.ini", FileMode.Create);
            StreamWriter SW = new StreamWriter(FS);

            SW.WriteLine("<CHANNEL 1>");
            SW.WriteLine("SCANNER1=" + scan1.Text);
            SW.WriteLine("");

            SW.Close();
            FS.Close();
        }
        public void loadconfig_device(ComboBox scan1)
        {
            string[] data = null;
            string str;
            FileStream FS = new FileStream(Application.StartupPath + @"\Device.ini", FileMode.Open);
            StreamReader SR = new StreamReader(FS);
            while (SR.EndOfStream == false)
            {
                str = SR.ReadLine();
                data = str.Split('=');

                switch (data[0])
                {
                    case "SCANNER1":
                        scan1.Text = data[1];
                        break;
                    
                    default:
                        break;
                }
            }

            SR.Close();
            FS.Close();
        }

        public void saveconfig_device(ComboBox scan1, ComboBox scan2, ComboBox scan3,
                                        ComboBox com1, ComboBox com2, ComboBox com3,
                                        ComboBox daq1, ComboBox daq2, ComboBox daq3,
                                        ComboBox dvm1, ComboBox dvm2, ComboBox dvm3,
                                        ComboBox hio1, ComboBox hio2, ComboBox hio3,
                                        ComboBox pwsl1, ComboBox pwsl2, ComboBox pwsl3,
                                        ComboBox sdiv1, ComboBox sdiv2, ComboBox sdiv3)
        {
            FileStream FS = new FileStream(Application.StartupPath + @"\Device.ini", FileMode.Create);
            StreamWriter SW = new StreamWriter(FS);

            SW.WriteLine("<CHANNEL 1>");
            SW.WriteLine("SCANNER1=" + scan1.Text);
            SW.WriteLine("COMBOARD1=" + com1.Text);
            SW.WriteLine("DAQ1=" + daq1.Text);
            SW.WriteLine("DVM1=" + dvm1.Text);
            SW.WriteLine("HIOKI1=" + hio1.Text);
            SW.WriteLine("POWERSUPPLY1=" + pwsl1.Text);
            SW.WriteLine("SDIVBOARD1=" + sdiv1.Text);
            SW.WriteLine("");

            //Thay đổi thêm save config kênh 2
            SW.WriteLine("<CHANNEL 2>");
            SW.WriteLine("SCANNER2=" + scan2.Text);
            SW.WriteLine("COMBOARD2=" + com2.Text);
            SW.WriteLine("DAQ2=" + daq2.Text);
            SW.WriteLine("DVM2=" + dvm2.Text);
            SW.WriteLine("HIOKI2=" + hio2.Text);
            SW.WriteLine("POWERSUPPLY2=" + pwsl2.Text);
            SW.WriteLine("SDIVBOARD2=" + sdiv2.Text);
            SW.WriteLine("");

            //Thay đổi thêm save config kênh 3
            SW.WriteLine("<CHANNEL 3>");
            SW.WriteLine("SCANNER3=" + scan3.Text);
            SW.WriteLine("COMBOARD3=" + com3.Text);
            SW.WriteLine("DAQ3=" + daq3.Text);
            SW.WriteLine("DVM3=" + dvm3.Text);
            SW.WriteLine("HIOKI3=" + hio3.Text);
            SW.WriteLine("POWERSUPPLY3=" + pwsl3.Text);
            SW.WriteLine("SDIVBOARD3=" + sdiv3.Text);
            SW.WriteLine("");

            SW.Close();
            FS.Close();
        }

        public void loadconfig_offset(TextBox tbocv1, TextBox tbir1,
                                        TextBox tbocv2, TextBox tbir2,
                                            TextBox tbocv3, TextBox tbir3)
        {
            string[] data = null;
            string str;
            FileStream FS = new FileStream(Application.StartupPath + @"\Offset.ini", FileMode.Open);
            StreamReader SR = new StreamReader(FS);
            while (SR.EndOfStream == false)
            {
                str = SR.ReadLine();
                data = str.Split('=');

                switch (data[0])
                {
                    case "OCV1":
                        tbocv1.Text = data[1];
                        break;
                    case "OCV2":
                        tbocv2.Text = data[1];
                        break;
                    case "OCV3":
                        tbocv3.Text = data[1];
                        break;
                    case "IR1":
                        tbir1.Text = data[1];
                        break;
                    case "IR2":
                        tbir2.Text = data[1];
                        break;
                    case "IR3":
                        tbir3.Text = data[1];
                        break;
                    default:
                        break;
                }
            }

            SR.Close();
            FS.Close();
        }

        public void saveconfig_offset(TextBox tbocv1, TextBox tbir1,
                                        TextBox tbocv2, TextBox tbir2,
                                           TextBox tbocv3, TextBox tbir3)
        {
            FileStream FS = new FileStream(Application.StartupPath + @"\Offset.ini", FileMode.Create);
            StreamWriter SW = new StreamWriter(FS);

            SW.WriteLine("<CHANNEL 1>");
            SW.WriteLine("OCV1=" + tbocv1.Text);
            SW.WriteLine("IR1=" + tbir1.Text);
            SW.WriteLine("");
            SW.WriteLine("<CHANNEL 2>");
            SW.WriteLine("OCV2=" + tbocv2.Text);
            SW.WriteLine("IR2=" + tbir2.Text);
            SW.WriteLine("");
            SW.WriteLine("<CHANNEL 3>");
            SW.WriteLine("OCV3=" + tbocv3.Text);
            SW.WriteLine("IR3=" + tbir3.Text);
            SW.WriteLine("");

            SW.Close();
            FS.Close();
        }

        public void loadconfig_Socket(TextBox lineno, TextBox mcid, TextBox stnid1, TextBox stnid2, TextBox stnid3, TextBox port, TextBox ip, TextBox workerid, TextBox socketport)
        {
            string[] data = null;
            string str;
            FileStream FS = new FileStream(Application.StartupPath + @"\Socket.ini", FileMode.Open);
            StreamReader SR = new StreamReader(FS);
            while (SR.EndOfStream == false)
            {
                str = SR.ReadLine();
                data = str.Split('=');

                switch (data[0])
                {
                    case "LINE_NO":
                        lineno.Text = data[1];
                        break;
                    case "MC_ID":
                        mcid.Text = data[1];
                        break;
                    case "STN_ID1":
                        stnid1.Text = data[1];
                        break;
                    case "STN_ID2":
                        stnid2.Text = data[1];
                        break;
                    case "STN_ID3":
                        stnid3.Text = data[1];
                        break;
                    case "PORT":
                        port.Text = data[1];
                        break;
                    case "IP":
                        ip.Text = data[1];
                        break;
                    case "ID":
                        workerid.Text = data[1];
                        break;
                    case "SKport":
                        socketport.Text = data[1];
                        break;
                    default:
                        break;
                }
            }
            SR.Close();
            FS.Close();
        }

        public void saveconfig_Socket(TextBox lineno, TextBox mcid, TextBox stnid1, TextBox stnid2, TextBox stnid3, TextBox port, TextBox ip, TextBox workerid, TextBox socketport)
        {
            FileStream FS = new FileStream(Application.StartupPath + @"\Socket.ini", FileMode.Create);
            StreamWriter SW = new StreamWriter(FS);

            SW.WriteLine("<Socket>");
            SW.WriteLine("LINE_NO=" + lineno.Text);
            SW.WriteLine("MC_ID=" + mcid.Text);
            SW.WriteLine("STN_ID1=" + stnid1.Text);
            SW.WriteLine("STN_ID2=" + stnid2.Text);
            SW.WriteLine("STN_ID3=" + stnid3.Text);
            SW.WriteLine("PORT=" + port.Text);
            SW.WriteLine("IP=" + ip.Text);
            SW.WriteLine("ID=" + workerid.Text);
            SW.WriteLine("SKport=" + socketport.Text);
            SW.WriteLine("");

            SW.Close();
            FS.Close();
        }

        public void loadconfig_Oracle(TextBox database, TextBox user, TextBox pass)
        {
            string[] data = null;
            string str;
            FileStream FS = new FileStream(Application.StartupPath + @"\Oracle.ini", FileMode.Open);
            StreamReader SR = new StreamReader(FS);
            while (SR.EndOfStream == false)
            {
                str = SR.ReadLine();
                data = str.Split('=');

                switch (data[0])
                {
                    case "DATABASE":
                        database.Text = data[1];
                        break;
                    case "USER":
                        user.Text = data[1];
                        break;
                    case "PASS":
                        pass.Text = data[1];
                        break;
                    default:
                        break;
                }
            }
            SR.Close();
            FS.Close();
        }

        public void saveconfig_Oracle(TextBox database, TextBox user, TextBox pass)
        {
            FileStream FS = new FileStream(Application.StartupPath + @"\Oracle.ini", FileMode.Create);
            StreamWriter SW = new StreamWriter(FS);

            SW.WriteLine("<Oracle>");
            SW.WriteLine("DATABASE=" + database.Text);
            SW.WriteLine("USER=" + user.Text);
            SW.WriteLine("PASS=" + pass.Text);
            SW.WriteLine("");

            SW.Close();
            FS.Close();
        }

        public string loadprocess()
        {
            try
            {
                string str = "";
                FileStream FS = new FileStream(Application.StartupPath + @"\Process.ini", FileMode.Open);
                StreamReader SR = new StreamReader(FS);
                while (SR.EndOfStream == false)
                {
                    str = SR.ReadLine();
                }
                SR.Close();
                FS.Close();
                return str;
            }
            catch (Exception)
            {
                return "CP-PT";
            }
        }

        public string loadprocess(ComboBox process)
        {
            try
            {
                string str = "";
                FileStream FS = new FileStream(Application.StartupPath + @"\Process.ini", FileMode.Open);
                StreamReader SR = new StreamReader(FS);
                while (SR.EndOfStream == false)
                {
                    str = SR.ReadLine();
                }
                SR.Close();
                FS.Close();
                process.Text = str;
                return str;
            }
            catch (Exception)
            {
                MessageBox.Show("file Processk ko có data");
               return "CP-PT";
            }
        }
        public void saveprocess(string process)
        {
            try
            {
                FileStream FS = new FileStream(Application.StartupPath + @"\Process.ini", FileMode.Create);
                StreamWriter SW = new StreamWriter(FS);
                SW.WriteLine(process);
                SW.Close();
                FS.Close();
            }
            catch (Exception)
            {

            }

        }

        public string loadlastmodel()
        {
            try
            {
                string str = "";
                FileStream FS = new FileStream(Application.StartupPath + @"\Lastmodel.ini", FileMode.Open);
                StreamReader SR = new StreamReader(FS);
                while (SR.EndOfStream == false)
                {
                    str = SR.ReadLine();
                }
                SR.Close();
                FS.Close();
                return str;
            }
            catch (Exception)
            {
                return "";
            }
        }

        public void savemodelname(ComboBox cbmodel)
        {
            try
            {
                FileStream FS = new FileStream(Application.StartupPath + @"\Lastmodel.ini", FileMode.Create);
                StreamWriter SW = new StreamWriter(FS);

                SW.WriteLine(cbmodel.Text);

                SW.Close();
                FS.Close();
            }
            catch (Exception)
            {
                
            }
            
        }

        public void loadqty(Label lbok, Label lbNG, Label lbtotal)
        {
            try
            {
                string[] data = null;
                string str;
                FileStream FS = new FileStream(Application.StartupPath + @"\Count.ini", FileMode.Open);
                StreamReader SR = new StreamReader(FS);
                while (SR.EndOfStream == false)
                {
                    str = SR.ReadLine();
                    data = str.Split('=');

                    switch (data[0])
                    {
                        case "OK":
                            lbok.Text = data[1];
                            break;
                        case "NG":
                            lbNG.Text = data[1];
                            break;
                        case "Total":
                            lbtotal.Text = data[1];
                            break;
                        default:
                            break;
                    }
                }
                SR.Close();
                FS.Close();
            }
            catch (Exception)
            {
                lbok.Text = "0";
                lbNG.Text = "0";
                lbtotal.Text = "0";
            }
        }
        public void loadqty(Label lbok1, Label lbNG1,Label lbok2,Label lbNG2,Label lbok3,Label lbNG3,ref int count,ref string ngay)
        {
            try
            {
                string[] data = null;
                string str;
                FileStream FS = new FileStream(Application.StartupPath + @"\Count.ini", FileMode.Open);
                StreamReader SR = new StreamReader(FS);
                while (SR.EndOfStream == false)
                {
                    str = SR.ReadLine();
                    data = str.Split('=');

                    switch (data[0])
                    {
                        case "OK1":
                            lbok1.Text = data[1];
                            break;
                        case "NG1":
                            lbNG1.Text = data[1];
                            break;
                        case "OK2":
                            lbok2.Text = data[1];
                            break;
                        case "NG2":
                            lbNG2.Text = data[1];
                            break;
                        case "OK3":
                            lbok3.Text = data[1];
                            break;
                        case "NG3":
                            lbNG3.Text = data[1];
                            break;
                        case "ngay":
                            ngay = data[1];
                            break;
                        case "count":
                            count = int.Parse(data[1]);
                            break;
                        default:
                            break;
                    }
                }
                SR.Close();
                FS.Close();
            }
            catch (Exception)
            {
                lbok1.Text = "0";
                lbNG1.Text = "0";
                lbok2.Text = "0";
                lbNG2.Text = "0";
            }
        }

        public void loadlog(ref int count, ref string ngay)
        {
            try
            {
                string[] data = null;
                string str;
                FileStream FS = new FileStream(Application.StartupPath + @"\Count.ini", FileMode.Open);
                StreamReader SR = new StreamReader(FS);
                while (SR.EndOfStream == false)
                {
                    str = SR.ReadLine();
                    data = str.Split('=');

                    switch (data[0])
                    {
                        case "OK1":
                            
                            break;
                        case "NG1":
                           
                            break;
                        case "OK2":
                            
                            break;
                        case "NG2":
                            
                            break;
                        case "OK3":

                            break;
                        case "NG3":

                            break;
                        case "ngay":
                            ngay = data[1];
                            break;
                        case "count":
                            count = int.Parse(data[1]);
                            break;
                        default:
                            break;
                    }
                }
                SR.Close();
                FS.Close();
            }
            catch (Exception)
            {
                
            }
        }

        public void saveqty(Label lbok, Label lbNG, Label lbtotal)
        {
            try
            {
                FileStream FS = new FileStream(Application.StartupPath + @"\Count.ini", FileMode.Create);
                StreamWriter SW = new StreamWriter(FS);

                SW.WriteLine("<Quantity>");
                SW.WriteLine("OK=" + lbok.Text);
                SW.WriteLine("NG=" + lbNG.Text);
                SW.WriteLine("Total=" + lbtotal.Text);

                SW.Close();
                FS.Close();
            }
            catch (Exception)
            {

            }

        }

        public void saveqty(Label lbok1, Label lbNG1,Label lbok2,Label lbNG2,Label lbok3,Label lbNG3,int count,string ngay)
        {
            try
            {
                FileStream FS = new FileStream(Application.StartupPath + @"\Count.ini", FileMode.Create);
                StreamWriter SW = new StreamWriter(FS);

                SW.WriteLine("<CH1>");
                SW.WriteLine("OK1=" + lbok1.Text);
                SW.WriteLine("NG1=" + lbNG1.Text);
                SW.WriteLine("<CH2>");
                SW.WriteLine("OK2=" + lbok2.Text);
                SW.WriteLine("NG2=" + lbNG2.Text);
                SW.WriteLine("<CH3>");
                SW.WriteLine("OK3=" + lbok3.Text);
                SW.WriteLine("NG3=" + lbNG3.Text);
                SW.WriteLine("ngay=" + ngay);
                SW.WriteLine("count=" + count);
                SW.Close();
                FS.Close();
            }
            catch (Exception)
            {

            }

        }

       
    }

}
