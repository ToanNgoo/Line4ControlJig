using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Threading;
using System.Data.OleDb;
using excell = Microsoft.Office.Interop.Excel;
using System.IO;
using Excel;
using System.IO.Ports;
using System.Runtime.InteropServices;
using System.Net.NetworkInformation;
namespace WindowsFormsApplication1
{
    public partial class Form1 : Form
    {
        // Adress PC connect with PLC
        //* PC sent PLC *//
        // Ready   -  "REDY" (M100)
        // RUN     -  "RUN!" (M101)
        // END     -  "END!" (M102)
        // END     -  "END!" (M1002)
        // Error   -  "EROR" (M103)
        // SPL END -  "SEND" (M104)
        // Manual  - M107
        // Start auto - M108
        // Stop auto - M111
        // Reset     - M112
        // Buzzer stop - M113
        // Up   - M109
        // Down - M110

        //* PLC sent PC* //
        // START TEST  -  "STRT" (M105)
        // TEST COMLETE    -  "STOP" (M106)
        // Flag on start test  - M205
        //#########################################################################################################

        private Thread saveLog;
        private Thread CheckPLC;
        private Thread checkServer;
        private Thread update;
        private Thread xuLiCodeJig;


        public int count1;
        clsCOM_MES com_pc;
        clsConfig config;
        clsPLC PLC;
        clsLocaldb localdb;
        clsLocaldb localdb1;
        clsScanner scaner;
        ClsExcel excel = new ClsExcel();

        public int checkCodeMaskboard = 0;
        public int dem = 0;
        public bool frm2 = true;
        bool kq = true;
        bool kq1 = true;
        bool trangThai = false;

        string status_Jig = string.Empty;
        string _model = string.Empty;
        string _modelLog = string.Empty;
        public string _modelCode = string.Empty;
        string _modelName = string.Empty;
        public string _numJig = string.Empty;

        public bool _checkServer = false;
        bool check48K = false;
        bool check49K = false;

        public string _dem = "";
        public string _demLocal = "";
        public string codeJig = "";
        public string code_MaskBoard = "";
        public Form1()
        {
            InitializeComponent();
            CheckForIllegalCrossThreadCalls = false;
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            tb_Ethernet.Text = DateTime.Now.ToString();
            btn_start.Enabled = false;
            com_pc = new clsCOM_MES();
            config = new clsConfig();
            PLC = new clsPLC();
            config.loadconfig_setup(tb_IPserver, tb_Link1, tb_Link2, tb_Link3, tb_Link4, tb_Link5, tb_Link6, tb_Ethernet, cbx_Scaner);
            init();
            status_Jig = PLC.readplc("X7");
            if (status_Jig == "1")
            {
                lbl_StatusJig.BackColor = Color.Green;
            }
            else
            {
                lbl_StatusJig.BackColor = Color.Red;
            }

            try
            {
                PLC.Writeplc("D100", int.Parse(tb_Value.Text));
                if (PLC.readplc("X16") == "1")
                {
                    lbl_StatusAir.BackColor = Color.Green;
                }
                else
                {
                    lbl_StatusAir.BackColor = Color.Red;
                }
            }

            catch (Exception)
            {
                ;
            }
            localdb = new clsLocaldb(tb_Link3.Text);
            localdb1 = new clsLocaldb(@Application.StartupPath);

            checkServer = new Thread(new ThreadStart(ping_Server));
            checkServer.IsBackground = true;
            checkServer.Start();

            scaner = new clsScanner(this);
            if (cbx_Scaner.Text != "")
            {
                scaner.COMnum = cbx_Scaner.Text;
                scaner.ketnoi(lbl_ScannerJig);
                scaner.ReadData();
                scaner.ReadData();
            }
            else
                MessageBox.Show("Chưa có thông tin COM Scanner", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);

            enable(tb_Ethernet.Text);
            timer2.Enabled = true;
            timer1.Enabled = true;
        }

        private void init()
        {
            txt_Cpu_type.Text = "520";
            txt_port_number.Text = "1";
            txt_IP_adress.Text = "192.168.1.10";
            txt_time_out.Text = "6000";
            txt_W_write.Text = "Adress";
            txt_W_value.Text = "Value";
            txt_R_adress.Text = "Adress";
            txt_R_result.Enabled = false;
            cbx_Com_PC.Text = "COM30";
            cbx_baudrate.Text = "19200";
            ketnoiCOM_MES(cbx_Com_PC, lbl_com_PC, com_pc);
            count1 = 0;
            PLC.thietlap(int.Parse(txt_Cpu_type.Text), int.Parse(txt_port_number.Text), txt_IP_adress.Text, int.Parse(txt_time_out.Text));
            PLC.ketnoi(lbl_PLC);
        }

        private void cbx_Com_PC_Click(object sender, EventArgs e)
        {
            config.loadlistcom(cbx_Com_PC);
        }

        private void btn_Knoi_COMPC_Click(object sender, EventArgs e)
        {
            if (cbx_baudrate.Text != "")
            {
                ketnoiCOM_MES(cbx_Com_PC, lbl_com_PC, com_pc);
            }
            else
            {
                MessageBox.Show("Hãy chọn Baudrate, sau đó chọn Line sử dụng", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }

        }

        private void ketnoiCOM_MES(ComboBox cb, Label lb, clsCOM_MES com)
        {
            switch (lb.BackColor.ToString())
            {
                case "Color [Red]":
                    com.COMnum = cb.Text;

                    if (com.ketnoi(cbx_baudrate.Text))
                    {
                        lb.BackColor = Color.Blue;
                        //saveconfigdevice();
                    }
                    break;
                case "Color [Blue]":
                    com.ngatketnoi();
                    lb.BackColor = Color.Red;
                    break;
                default:
                    break;
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {

        }





        private void timer1_Tick(object sender, EventArgs e)
        {
            //label58.Text = PLC.readplc("M105");

            try
            {
                if (PLC.readplc("X6") == "1" && count1 == 1)
                {
                    PLC.Writeplc("M105", 0);
                    count1 = 0;
                }

                if (PLC.readplc("M105") == "1" && count1 == 0)
                {
                    count1 = 1;
                    if (using_QR)
                    {
                        if (!checkBox1.Checked)
                        {
                            xuLiCodeJig = new Thread(new ThreadStart(ReadCodeMaskBoard));
                            xuLiCodeJig.IsBackground = true;
                            xuLiCodeJig.Start();

                            if (tb_Link4.BackColor == Color.Red)
                            {
                                DialogResult rs = MessageBox.Show("Mất kết nối đến hệ thống QR code, kiểm tra lại kết nối theo đường Link 4", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            }
                            else
                            {
                                saveLog = new Thread(new ThreadStart(Input_Inspection));
                                saveLog.IsBackground = true;
                                saveLog.Start();
                            }
                        }
                        else
                        {
                            com_pc.sent_data("STRT");
                            PLC.Writeplc("M205", 1);
                            Thread.Sleep(300);
                            PLC.Writeplc("M205", 0);
                            count1 = 0;
                        }
                    }
                    else
                    {
                        com_pc.sent_data("STRT");
                        PLC.Writeplc("M205", 1);
                        Thread.Sleep(300);
                        PLC.Writeplc("M205", 0);
                        count1 = 0;
                    }

                }
            }
            catch (Exception error)
            {
                MessageBox.Show("Lỗi Auto Start Test :\r\n" + error.ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }


            try
            {
                if (com_pc.Data == "END!" || com_pc.UpJig == "UP")
                {
                    com_pc.Data = "";
                    com_pc.UpJig = "";
                    PLC.Writeplc("M102", 1);
                    Thread.Sleep(300);
                    PLC.Writeplc("M102", 0);

                    //if (using_QR && !checkBox1.Checked)
                    //{
                    //    if (tb_Link4.BackColor == Color.Red)
                    //    {
                    //        DialogResult rs = MessageBox.Show("Mất kết nối đến hệ thống QR code, kiểm tra lại kết nối theo đường Link 4", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    //    }
                    //    else
                    //    {
                    //        saveLog = new Thread(new ThreadStart(Input_Inspection));
                    //        saveLog.IsBackground = true;
                    //        saveLog.Start();
                    //    }
                    //}

                    if (dem > 50000)
                    {
                        MessageBox.Show("Đã hết hạn sử dụng kim");
                        kq1 = false;
                        PLC.Writeplc("M108", 0);
                        PLC.Writeplc("M111", 1);
                    }
                    else
                    {
                        update = new Thread(new ThreadStart(up_server));
                        update.IsBackground = true;
                        update.Start();
                    }
                }
            }
            catch (Exception error)
            {
                MessageBox.Show("Lỗi Auto Up Jig :\r\n" + error.ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            if (com_pc.Data == "READ")
            {
                com_pc.Data = "";
                PLC.Writeplc("M100", 1);
                Thread.Sleep(300);
                PLC.Writeplc("M100", 0);
            }

            if (PLC.readplc("M106") == "1" && count1 == 0)
            {
                count1 = 1;
                com_pc.sent_data("STOP");
                count1 = 0;
            }

        }

        private void btn_manual_Click(object sender, EventArgs e)
        {
            PLC.Writeplc("M108", 0);
            PLC.Writeplc("M107", 1);
            PLC.Writeplc("M112", 0);
            btn_Up.Enabled = true;
            btn_down.Enabled = true;
        }

        private void btn_Up_Click(object sender, EventArgs e)
        {
            PLC.Writeplc("M110", 0);
            PLC.Writeplc("M109", 1);
        }

        private void btn_down_Click(object sender, EventArgs e)
        {
            PLC.Writeplc("M109", 0);
            PLC.Writeplc("M110", 1);
        }

        private void btn_write_Click(object sender, EventArgs e)
        {
            PLC.Writeplc(txt_W_write.Text, int.Parse(txt_W_value.Text));
        }

        private void btn_read_Click(object sender, EventArgs e)
        {
            txt_R_result.Text = PLC.readplc(txt_R_adress.Text);
        }

        private void txt_W_write_Click(object sender, EventArgs e)
        {
            txt_W_write.Text = "";
        }
        private void txt_R_adress_Click(object sender, EventArgs e)
        {
            txt_R_adress.Text = "";
        }

        private void txt_W_value_Click(object sender, EventArgs e)
        {
            txt_W_value.Text = "";
        }

        private void btn_sent_UART_Click(object sender, EventArgs e)
        {
            com_pc.sent_data_byte(txt_factor1.Text, txt_factor2.Text, txt_factor3.Text, txt_factor4.Text);
        }

        private void button2_Click(object sender, EventArgs e)
        {
            bool Isopen = false;
            if (tb_barcode.Text == "")
            {
                MessageBox.Show("Bạn chưa khởi động chương trình đếm pin Jig và auto check thông số, Hãy khởi động!");
            }
            else
            {
                foreach (Form f in Application.OpenForms)
                {
                    if (f.Text == "Form2")
                    {
                        Isopen = true;
                        f.BringToFront();
                        break;
                    }
                }
                if (Isopen == false)
                {
                    Form2 mt = new Form2(tb_Link3.Text, @Application.StartupPath, this);
                    mt.TopMost = true;
                    mt.Show();
                }
                btn_down.Enabled = false;
                btn_Up.Enabled = false;
                timer1.Enabled = true;
                PLC.Writeplc("M107", 0);
                if (kq == true && kq1 == true && PLC.readplc("X16") == "1")
                {
                    PLC.Writeplc("M111", 0);
                    PLC.Writeplc("M108", 1);
                }
                else if (kq == false || kq1 == false || PLC.readplc("X16") == "0")
                {
                    PLC.Writeplc("M111", 1);
                    PLC.Writeplc("M108", 0);
                    if (PLC.readplc("X16") == "0")
                    {
                        MessageBox.Show(" Áp lực khí không đủ");
                    }
                }
                PLC.Writeplc("M107", 0);
                PLC.Writeplc("M112", 0);


                try
                {
                    if (PLC.readplc("X16") == "1")
                    {
                        lbl_StatusAir.BackColor = Color.Green;
                    }
                    else
                    {
                        lbl_StatusAir.BackColor = Color.Red;
                    }
                }
                catch (Exception)
                {
                    ;
                }
            }
        }

        private void btn_stop_Click(object sender, EventArgs e)
        {
            {
                btn_down.Enabled = false;
                btn_Up.Enabled = false;
                timer1.Enabled = true;
                PLC.Writeplc("M107", 0);
                PLC.Writeplc("M108", 0);
                PLC.Writeplc("M111", 1);
                PLC.Writeplc("M112", 0);
            }
        }

        private void btn_reset_Click(object sender, EventArgs e)
        {
            {
                btn_down.Enabled = false;
                btn_Up.Enabled = false;
                timer1.Enabled = true;
                PLC.Writeplc("M107", 0);
                PLC.Writeplc("M108", 0);
                PLC.Writeplc("M111", 0);
                PLC.Writeplc("M112", 1);
            }
        }

        public string code_Jig = "";
        bool using_QR = false;
        private void btn_kt_Click(object sender, EventArgs e)
        {
            if (tb_barcode.Text != "")
            {
                code_Jig = tb_barcode.Text;

                if (tb_Link1.BackColor == Color.Red || tb_Link2.BackColor == Color.Red)
                {
                    if (tb_Link1.BackColor == Color.Red)
                    {
                        DialogResult rs = MessageBox.Show("Đường link đến Folder chứa File master không đúng, Kiểm tra lại Link 1", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }

                    if (tb_Link2.BackColor == Color.Red)
                    {
                        DialogResult rs = MessageBox.Show("Đường link đến Folder chứa Log File không đúng, Kiểm tra lại Link 2", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
                else
                {
                    dem = 0;
                    dem_Local = 0;
                    callModel(tb_barcode.Text);
                }
            }
            else
            {
                DialogResult rs = MessageBox.Show("Chưa nhập code Jig", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        string FCT_Program = "";
        public void callModel(string codeJig)
        {
            try
            {
                if (tb_Link3.BackColor == Color.Red)
                {
                    FCT_Program = localdb1._getProgram(codeJig);
                    using_QR = localdb1.Status_QR(codeJig);
                }
                else
                {
                    FCT_Program = localdb._getProgram(codeJig);
                    using_QR = localdb.Status_QR(codeJig);
                }

                if (FCT_Program == "" || FCT_Program == null)
                {
                    DialogResult rs = MessageBox.Show("Code Jig bạn nhập không tồn tại trên hệ thống", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                else
                {
                    dem_Local = localdb1.loadCounting(codeJig);
                    if (tb_Link3.BackColor == Color.Green)
                    {
                        dem = localdb.loadCounting(codeJig);
                        if (dem_Local > 0)
                        {
                            localdb.uploadCounting(codeJig, dem + dem_Local);
                            localdb1.uploadCounting(codeJig, 0);
                            dem = localdb.loadCounting(codeJig);
                            _dem = dem.ToString();
                        }
                        else if (dem_Local == 0)
                        {
                            dem = localdb.loadCounting(codeJig);
                            _dem = dem.ToString();
                        }
                    }
                    else
                    {
                        dem_Local = localdb1.loadCounting(codeJig);
                        _demLocal = dem_Local.ToString();
                    }

                    string linkMaster = string.Empty;
                    linkMaster = tb_Link1.Text;
                    string linkLog = string.Empty;
                    linkLog = tb_Link2.Text; // link log File
                    string _File = getFile(@linkLog);

                    if (ReadFCTProgram(_File) == "" && !OpenFile)
                    {
                        MessageBox.Show("File Log Function mới nhất đang mở\r\nTắt Log File trước khi kiểm tra", "Lỗi mở File", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                    else
                    {
                        if (ReadFCTProgram(_File).Contains(FCT_Program))
                        {
                            int check = 0;
                            string[] list = Directory.GetFiles(@linkMaster);
                            foreach (string tmp in list)
                            {
                                if (tmp == @linkMaster + "\\" + FCT_Program + ".csv")
                                {
                                    check++;
                                }
                            }

                            if (check > 0)
                            {
                                kq = kiemTra(@linkMaster + "\\" + FCT_Program + ".csv", _File);
                                if (kq == false)
                                {
                                    btn_start.Enabled = false;
                                    PLC.Writeplc("M111", 1);
                                    PLC.Writeplc("M108", 0);
                                    lbl_Parameter.BackColor = Color.Red;
                                    DialogResult rsl = MessageBox.Show("Kết quả kiểm tra NG" +
                                                                        "\r\nNguyên nhân 1 : Chưa test bản function đầu tiên trước khi kiểm tra" +
                                                                        "\r\nNguyên nhân 2 : File Master và File Log Khác nhau ", "Thông báo", MessageBoxButtons.OKCancel, MessageBoxIcon.Error);

                                }
                                else
                                {
                                    btn_start.Enabled = true;
                                    PLC.Writeplc("M111", 0);
                                    lbl_Parameter.BackColor = Color.Green;
                                }
                            }
                            else
                            {
                                DialogResult rsl = MessageBox.Show("Kết quả kiểm tra NG" +
                                                                            "\r\nNguyên nhân : Không tìm thấy file master"
                                                                            , "Thông báo", MessageBoxButtons.OKCancel, MessageBoxIcon.Error);
                            }

                        }
                        else
                        {
                            btn_start.Enabled = false;
                            PLC.Writeplc("M111", 1);
                            PLC.Writeplc("M108", 0);
                            lbl_Parameter.BackColor = Color.Red;
                            DialogResult rsl = MessageBox.Show("Kết quả kiểm tra NG" +
                                                               "\r\nNguyên nhân 1 : Chưa test bản function đầu tiên trước khi kiểm tra" +
                                                               "\r\nNguyên nhân 2 : File Master và File Log Khác nhau ", "Thông báo", MessageBoxButtons.OKCancel, MessageBoxIcon.Error);
                        }
                    }

                }
            }
            catch (Exception ee)
            {
                DialogResult rs = MessageBox.Show(ee.ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            trangThai = false;
        }

        public string getFile(string link)
        {
            string File_OK = "";
            List<string> Folder = new List<string>();
            List<string> File_Final = new List<string>();

            string[] listModelFCT = localdb1.loadProgramFCT(tb_barcode.Text);        

            if (listModelFCT.Length > 0)
            {
                for (int i = 0; i < listModelFCT.Length; i++)
                { 
                    if(Directory.Exists(@tb_Link2.Text + @"\" + listModelFCT[i] + @"\" + DateTime.Now.Year.ToString() + "-" + DateTime.Now.Month.ToString("00")))
                    {
                        Folder.Add(@tb_Link2.Text + @"\" + listModelFCT[i] + @"\" + DateTime.Now.Year.ToString() + "-" + DateTime.Now.Month.ToString("00"));
                    }
                }
            }

            foreach (string l in Folder)
            {
                File_Final.Add(SerchFile(l));                
            }


            int pos = 0;
            File_OK = File_Final[pos];

            for (int i = 1; i < File_Final.Count; i++)
            {
                if (File.GetLastWriteTime(@File_Final[i]) > File.GetLastWriteTime(@File_OK))
                {
                    File_OK = File_Final[i];                
                }
            }
            return File_OK;          
        }


        public string SerchFile(string link)
        {
            string[] file = Directory.GetFiles(@link); // link đến vị trí fileLog

            DateTime[] dt = new DateTime[file.Length];
            for (int i = 0; i < file.Length; i++)
            {
                dt[i] = File.GetLastWriteTime(file[i]);
            }
            DateTime maxTime = dt[0];
            int vTri = 0;
            for (int i = 1; i < dt.Length; i++)
            {
                int soSanh = DateTime.Compare(dt[i], maxTime);
                if (soSanh > 0)
                {
                    maxTime = dt[i];
                    vTri = i;
                }
            }
            return file[vTri];
        }


        public bool OpenFile = true;
        public bool kiemTra(string linkFile1, string linkFile2)
        {
            ////khoi tao table1
            DataTable tbl1 = ReadCsvFile(@linkFile1);
            dtv_tb1.DataSource = tbl1;

            // khoi tao table2
            DataTable tbl2 = ReadCsvFile(@linkFile2);
            dtv_tb2.DataSource = tbl2;

            bool check = SosanhTable(tbl1, tbl2);
            return check;
        }

        public string ReadFCTProgram(string path)
        {
            OpenFile = true;
            int i = 0;
            int counting = 0;
            string FCTprogram = "";
            string Fulltext = "";
        aaa:
            try
            {
                using (StreamReader sr = new StreamReader(path))
                {
                    while (!sr.EndOfStream)
                    {
                        Fulltext = sr.ReadToEnd().ToString();
                    }
                }
                i = 0;
            }
            catch (Exception)
            {
                i = 1;
                counting++;
            }

            if (i == 1 && counting < 10)
                goto aaa;
            else
            {
                if (i == 0)
                {
                    OpenFile = true;
                }
                else
                {
                    OpenFile = false;
                }
            }


            if (OpenFile)
            {
                string[] rows = Fulltext.Split('\r');
                string[] value = rows[0].Split(',');
                FCTprogram = value[0];
                return FCTprogram;
            }
            else
            {
                return "";
            }
        }


        public DataTable ReadCsvFile(string path)
        {
            DataTable dtb = new DataTable();
            string Fulltext;
            using (StreamReader sr = new StreamReader(path))
            {
                while (!sr.EndOfStream)
                {
                    Fulltext = sr.ReadToEnd().ToString();
                    string[] rows = Fulltext.Split('\r');

                    for (int i = 7; i < 10; i++)
                    {
                        string[] rowValues = rows[i].Split(',');

                        if (i == 7)
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
            }
            return dtb;
        }

        private void bt_REset_Click(object sender, EventArgs e)
        {
            PLC.Writeplc("M3500", 0);
        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            if (lbl_Parameter.BackColor == Color.Green)
            {
                string user = "";
                string pass = "";
                config.loaduser(user, pass);
                if (user == tb_CPE.Text && pass == tb_passCPE.Text)
                {
                    if (_checkServer)
                    {
                        localdb.uploadCounting(code_Jig, 0);
                        localdb1.uploadCounting(code_Jig, 0);
                        tb_CPE.Text = "";
                        tb_passCPE.Text = "";
                        DialogResult rs = MessageBox.Show("Reset thành công", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }
                    else
                    {
                        DialogResult rs = MessageBox.Show("Không kết nối đến server, không thể reset", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }
                }
                else
                {
                    DialogResult rs = MessageBox.Show("Tài khoản bạn nhập không đúng, vui lòng nhập lại", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    tb_CPE.Text = "";
                    tb_passCPE.Text = "";
                }
            }
            else
            {
                MessageBox.Show("Hãy nhập code Jig và check thông số trước khi Reset");
            }
        }

        public bool IsNumber(string pValue)
        {
            foreach (Char c in pValue)
            {
                if (!Char.IsDigit(c))
                    return false;
            }
            return true;
        }        

        private void cb_BuzzStop_CheckedChanged(object sender, EventArgs e)
        {
            if (cb_BuzzStop.Checked == true)
            {
                PLC.Writeplc("M90", 1);
            }
            else
                PLC.Writeplc("M90", 0);
        }

        private void button2_Click_1(object sender, EventArgs e)
        {
            try
            {
                PLC.Writeplc("D100", int.Parse(tb_Value.Text));
                File.WriteAllText(@Application.StartupPath + "\\PLC.txt", tb_Value.Text);
            }
            catch (Exception)
            {
                ;
            }
        }

        public static bool SosanhTable(DataTable tbl1, DataTable tbl2)
        {
            if (tbl1.Rows.Count != tbl2.Rows.Count || tbl1.Columns.Count != tbl2.Columns.Count)
                return false;
            for (int i = 0; i < tbl1.Rows.Count; i++)
            {
                for (int c = 0; c < tbl1.Columns.Count; c++)
                {
                    if (!Equals(tbl1.Rows[i][c], tbl2.Rows[i][c]))
                        return false;
                }
            }
            return true;
        }

        public void Input_Inspection()
        {
            Thread.Sleep(1000);
            int vTri = 0;
            string link = tb_Link2.Text + "\\" + FCT_Program;
            string logFunction = link + "\\" + DateTime.Now.Year.ToString() + "-" + DateTime.Now.Month.ToString("00");
            string save = WriteFile();
            string saveLocal = WriteFileLocal();
            string[] listFile = Directory.GetFiles(@logFunction);
            DateTime[] dt = new DateTime[listFile.Length];
            for (int i = 0; i < listFile.Length; i++)
            {
                dt[i] = File.GetLastWriteTime(listFile[i]);
            }

            // Tìm File mới được ghi nhất
            DateTime time1 = dt[0];
            for (int i = 1; i < dt.Length; i++)
            {
                if (DateTime.Compare(dt[i], time1) > 0)
                {
                    time1 = dt[i];
                    vTri = i;
                }
            }

            DataTable table = new DataTable();

            table = excel.ReadLog(listFile[vTri]);
            DataTable table_1 = new DataTable();

            string[] code = new string[40];
            string[] result = new string[40];
            string[] stt = new string[40];
            string[] data = new string[40];
            int tmp = 39;
            if (table.Rows.Count > 40)
            {
                for (int i = table.Rows.Count - 2; i > table.Rows.Count - 42; i--)
                {
                    code[tmp] = table.Rows[i].ItemArray[2].ToString();
                    result[tmp] = table.Rows[i].ItemArray[1].ToString();
                    tmp--;
                }

                for (int i = 0; i < 40; i++)
                {
                    data[i] = code[i] + "-" + result[i];
                }
                File.WriteAllLines(save, data);
                File.WriteAllLines(saveLocal, data);
            }
            else
            {
                ;
            }
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            //_frm2.Abort();
            timer1.Enabled = false;            
            com_pc.ngatketnoi();
            //if (CheckPLC.IsAlive) CheckPLC.Abort();
            //if (checkServer.IsAlive) checkServer.Abort();
            dem = 0;
        }

        public string WriteFile()
        {
            string File = tb_Link4.Text;
            File += "\\" + DateTime.Now.Hour.ToString("00") + DateTime.Now.Minute.ToString("00") + DateTime.Now.Second.ToString("00") + ".txt";
            return File;
        }

        public string WriteFileLocal()
        {
            string File = tb_Link4.Text;
            File += "\\Old\\" + DateTime.Now.Hour.ToString("00") + DateTime.Now.Minute.ToString("00") + DateTime.Now.Second.ToString("00") + ".txt";
            return File;
        }

        public bool Ping_Server(string IP)
        {
            int tmp = 0;
            try
            {
                Ping myPing = new Ping();
                PingReply reply = myPing.Send(IP, 1000);
                for (int i = 0; i < 2; i++)
                {
                    reply = myPing.Send(IP, 1000);
                    if (reply.Status == IPStatus.Success)
                    {
                        tmp++;
                    }
                    else
                    {
                        tmp = 0;
                        break;
                    }
                }

                if (tmp >= 1)
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

        // bool dis = false;
        public int dem_Local = 0;

        public void up_server()
        {
            if (tb_Link3.BackColor == Color.Green)
            {
                try
                {
                    dem_Local = localdb1.loadCounting(code_Jig);
                    dem = localdb.loadCounting(code_Jig) + 1;
                    if (dem_Local > 0)
                    {
                        localdb.uploadCounting(code_Jig, dem + dem_Local);
                        localdb1.uploadCounting(code_Jig, 0);
                        dem = localdb.loadCounting(code_Jig) + 1;
                        _dem = dem.ToString();
                    }
                    else
                    {
                        localdb.uploadCounting(code_Jig, dem);
                        _dem = dem.ToString();
                    }
                }
                catch (Exception)
                {

                }
            }
            else
            {
                try
                {
                    dem_Local = localdb1.loadCounting(code_Jig) + 1;
                    localdb1.uploadCounting(code_Jig, dem_Local);
                    _demLocal = dem_Local.ToString();
                }
                catch (Exception)
                {

                }

            }
            frm2 = false;
        }

        bool sv = true;
        public void ping_Server()
        {
            try
            {
                while (true)
                {
                    if (sv)
                    {
                        sv = false;
                        if (Ping_Server(tb_IPserver.Text))
                        {
                            lbl_Server.BackColor = Color.Green;
                            tb_IPserver.BackColor = Color.Green;
                            btn_Disable.BackColor = Color.FromArgb(192, 255, 192);
                            _checkServer = true;
                        }
                        else
                        {
                            _checkServer = false;
                            tb_IPserver.BackColor = Color.Red;
                            lbl_Server.BackColor = Color.Red;
                            btn_Disable.BackColor = Color.Red;
                        }

                        try
                        {
                            if (Directory.Exists(@tb_Link1.Text))
                            {
                                tb_Link1.BackColor = Color.Green;
                            }
                            else
                            {
                                tb_Link1.BackColor = Color.Red;
                            }
                        }
                        catch (Exception)
                        {
                            tb_Link1.BackColor = Color.Red;
                        }

                        if (Directory.Exists(@tb_Link2.Text))
                        {
                            tb_Link2.BackColor = Color.Green;
                        }
                        else
                        {
                            tb_Link2.BackColor = Color.Red;
                        }

                        try
                        {
                            if (Directory.Exists(@tb_Link3.Text))
                            {
                                tb_Link3.BackColor = Color.Green;

                            }
                            else
                            {
                                tb_Link3.BackColor = Color.Red;

                            }
                        }
                        catch (Exception)
                        {
                            tb_Link3.BackColor = Color.Red;
                        }

                        try
                        {
                            if (Directory.Exists(@tb_Link4.Text))
                            {
                                tb_Link4.BackColor = Color.Green;
                            }
                            else
                            {
                                tb_Link4.BackColor = Color.Red;
                            }
                        }
                        catch (Exception)
                        {
                            tb_Link4.BackColor = Color.Red;
                        }

                        try
                        {
                            if (Directory.Exists(@tb_Link5.Text))
                            {
                                tb_Link5.BackColor = Color.Green;
                            }
                            else
                            {
                                tb_Link5.BackColor = Color.Red;
                            }
                        }
                        catch (Exception)
                        {
                            tb_Link5.BackColor = Color.Red;
                        }

                        try
                        {
                            if (Directory.Exists(@tb_Link6.Text))
                            {
                                tb_Link6.BackColor = Color.Green;
                            }
                            else
                            {
                                tb_Link6.BackColor = Color.Red;
                            }
                        }
                        catch (Exception)
                        {
                            tb_Link6.BackColor = Color.Red;
                        }

                        sv = true;
                    }
                }
            }
            catch (Exception)
            {
                //MessageBox.Show("lỗi kiểm tra server :\r\n" + error.ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                sv = true;
            }

        }

        bool _plc = true;


        int step = 1;
        public void PLC_status()
        {
            switch (step)
            {
                case 1:
                    if (PLC.readplc("X6") == "1")
                    {
                        lbl_sol_up.BackColor = Color.Green;
                    }
                    else
                    {
                        lbl_sol_up.BackColor = Color.White;
                    }
                    step = 2;
                    break;
                case 2:
                    if (PLC.readplc("X5") == "1")
                    {
                        lbl_sol_down.BackColor = Color.Green;
                    }
                    else
                    {
                        lbl_sol_down.BackColor = Color.White;
                    }
                    step = 3;
                    break;
                case 3:
                    if (PLC.readplc("Y0") == "1")
                    {
                        lbl_towerlamp_R.BackColor = Color.Red;
                    }
                    else
                    {
                        lbl_towerlamp_R.BackColor = Color.White;
                    }
                    step = 4;
                    break;
                case 4:
                    if (PLC.readplc("Y1") == "1")
                    {
                        lbl_towerlamp_Y.BackColor = Color.Yellow;
                    }
                    else
                    {
                        lbl_towerlamp_Y.BackColor = Color.White;
                    }
                    step = 5;
                    break;
                case 5:
                    if (PLC.readplc("Y2") == "1")
                    {
                        lbl_towerlamp_G.BackColor = Color.Green;
                    }
                    else
                    {
                        lbl_towerlamp_G.BackColor = Color.White;
                    }
                    step = 6;
                    break;
                case 6:
                    if (PLC.readplc("Y03") == "1")
                    {
                        lbl_towerlamp_B.BackColor = Color.Gray;
                    }
                    else
                    {
                        lbl_towerlamp_B.BackColor = Color.White;
                    }
                    step = 7;
                    break;
                case 7:
                    if (PLC.readplc("Y4") == "1")
                    {
                        lbl_sol_up.BackColor = Color.Green;
                    }
                    else
                    {
                        lbl_sol_up.BackColor = Color.White;
                    }
                    step = 8;
                    break;
                case 8:
                    if (PLC.readplc("Y5") == "1")
                    {
                        lbl_sol_down.BackColor = Color.Green;
                    }
                    else
                    {
                        lbl_sol_down.BackColor = Color.White;
                    }
                    step = 9;
                    break;
                case 9:
                    if (PLC.readplc("Y4") == "0")
                    {
                        lbl_sol_off.BackColor = Color.Red;
                    }
                    else
                    {
                        lbl_sol_off.BackColor = Color.White;
                    }
                    step = 10;
                    break;
                case 10:
                    if (PLC.readplc("Y5") == "0")
                    {
                        lbl_sol_off.BackColor = Color.Red;
                    }
                    else
                    {
                        lbl_sol_off.BackColor = Color.White;
                    }
                    step = 11;
                    break;
                case 11:
                    if (PLC.readplc("M1000") == "1")
                    {
                        btn_reset.BackColor = Color.Yellow;
                    }
                    else
                    {
                        btn_reset.BackColor = Color.Gray;
                    }
                    step = 12;
                    break;
                case 12:
                    if (PLC.readplc("M107") == "1")
                    {
                        btn_manual.BackColor = Color.Yellow;
                    }
                    else
                    {
                        btn_manual.BackColor = Color.Gray;
                    }
                    step = 13;
                    break;
                case 13:
                    if (PLC.readplc("M109") == "1")
                    {
                        btn_Up.BackColor = Color.Green;
                    }
                    else
                    {
                        btn_Up.BackColor = Color.Gray;
                    }
                    step = 14;
                    break;
                case 14:
                    if (PLC.readplc("M110") == "1")
                    {
                        btn_down.BackColor = Color.Green;
                    }
                    else
                    {
                        btn_down.BackColor = Color.Gray;
                    }
                    step = 1;
                    break;

                default:
                    break;

            }
            _plc = true;
        }

        private void button3_Click(object sender, EventArgs e)
        {

            string status = "";
            if (tb_jig.Text == "" || tb_FCT.Text == "" || (!rbt_NoUse.Checked && !rbt_Use.Checked))
            {
                DialogResult rs = MessageBox.Show("Bạn nhập thiếu thông tin", "Thông báo", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
            }
            else
            {
                DialogResult rs = MessageBox.Show("Bạn có chắc chắn thêm model này", "Thông báo", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                if (rs == DialogResult.Yes)
                {
                    if (rbt_NoUse.Checked)
                    {
                        status = "No";
                    }

                    if (rbt_Use.Checked)
                    {
                        status = "Yes";
                    }
                    if (_checkServer)
                    {
                        localdb.uploadmodel(tb_jig.Text, tb_FCT.Text, status);
                        localdb1.uploadmodel(tb_jig.Text, tb_FCT.Text, status);
                    }
                    else
                    {
                        DialogResult rst = MessageBox.Show("Server không kết nối, không thể thêm model", "Thông báo", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                    }
                }
            }
        }

        public void checkModel()
        {
            int compare = 0;
            string[] model_Local = localdb1.getModel();
            string[] model_Server = localdb.getModel();
            for (int i = 0; i < model_Local.Length; i++)
            {
                for (int j = 0; j < model_Server.Length; j++)
                {
                    if (model_Local[i] == model_Server[j])
                    {
                        compare++;
                    }
                }
            }
        }

        public void readCodeJig()
        {
            scaner.ReadData();
        }

        static void enable(string name)
        {
            System.Diagnostics.ProcessStartInfo psi = new System.Diagnostics.ProcessStartInfo("netsh", "interface set interface \"" + name + "\" enable");
            System.Diagnostics.Process p = new System.Diagnostics.Process();
            p.StartInfo = psi;
            p.Start();
        }

        static void disable(string name)
        {
            System.Diagnostics.ProcessStartInfo psi = new System.Diagnostics.ProcessStartInfo("netsh", "interface set interface \"" + name + "\" disable");
            System.Diagnostics.Process p = new System.Diagnostics.Process();
            p.StartInfo = psi;
            p.Start();
        }

        private void button5_Click_1(object sender, EventArgs e)
        {
            if (btn_Disable.Text == "Disable")
            {
                disable(tb_Ethernet.Text);
                try
                {
                    if (checkServer.IsAlive == true)
                        checkServer.Abort();
                    btn_Disable.Text = "Enable";
                    //tb_Link1.BackColor = Color.Red;
                    //tb_Link2.BackColor = Color.Red;
                    tb_Link3.BackColor = Color.Red;
                    tb_Link4.BackColor = Color.Red;
                    tb_Link5.BackColor = Color.Red;
                    tb_Link6.BackColor = Color.Red;
                    btn_Disable.BackColor = Color.Red;
                    lbl_Server.BackColor = Color.Red;
                }
                catch (Exception error)
                {
                    MessageBox.Show("Lỗi khi ngắt kết nối Internet :\r\n" + error.ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            else
            {
                enable(tb_Ethernet.Text);
                btn_Disable.Text = "Disable";
                checkServer = new Thread(new ThreadStart(ping_Server));
                checkServer.IsBackground = true;
                checkServer.Start();
            }

        }

        private void button4_Click_1(object sender, EventArgs e)
        {
            enable(tb_Ethernet.Text);
        }


        bool checkInShift = true;
        private void timer2_Tick(object sender, EventArgs e)
        {
            if (_plc)
            {
                _plc = false;
                CheckPLC = new Thread(new ThreadStart(PLC_status));
                CheckPLC.IsBackground = true;
                CheckPLC.Start();
            }



            if (dem > 48000)
            {
                if (dem > 48000 && dem < 49000 && check48K == false)
                {
                    check48K = true;
                    PLC.Writeplc("M113", 1);
                    DialogResult result = new System.Windows.Forms.DialogResult();
                    result = MessageBox.Show("Sắp đến chu kỳ thay kim hãy báo ME, PE", "question", MessageBoxButtons.OK);
                    if (result == DialogResult.OK)
                    {
                        PLC.Writeplc("M113", 0);
                    }
                }
                else if (dem > 49000 && dem < 50000 && check49K == false)
                {
                    check49K = true;
                    PLC.Writeplc("M113", 1);
                    DialogResult result = new System.Windows.Forms.DialogResult();
                    result = MessageBox.Show("Sắp đến chu kỳ thay kim hãy báo ME, PE", "question", MessageBoxButtons.OK);
                    if (result == DialogResult.Yes)
                    {
                        PLC.Writeplc("M113", 0);
                    }
                }
            }


            if (PLC.readplc("M50") == "1")
            {
                btn_start.BackColor = Color.Blue;
                btn_stop.BackColor = Color.Gray;
            }
            else
            {
                btn_start.BackColor = Color.Gray;
                btn_stop.BackColor = Color.Red;
            }

            if (DateTime.Now.Second < 5 && DateTime.Now.Minute % 5 == 0 && trangThai == false)
            {
                checkInShift = false;
                trangThai = true;
                btn_kt.PerformClick();
                checkInShift = true;
            }
        }


        private void cbx_Scaner_Click(object sender, EventArgs e)
        {
            cbx_Scaner.Items.Clear();
            cbx_Scaner.Items.AddRange(SerialPort.GetPortNames());
        }

        private void btn_KetNoi_Click(object sender, EventArgs e)
        {
            try
            {
                if (lbl_ScannerJig.BackColor == Color.Green)
                {
                    scaner.ngatketnoi();
                    if (cbx_Scaner.Text != "")
                    {
                        scaner.COMnum = cbx_Scaner.Text;
                        scaner.ketnoi(lbl_ScannerJig);
                        scaner.ReadData();
                        scaner.ReadData();
                    }
                    else
                        MessageBox.Show("Chưa có thông tin COM Scanner", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                else
                {
                    if (cbx_Scaner.Text != "")
                    {
                        scaner.COMnum = cbx_Scaner.Text;
                        scaner.ketnoi(lbl_ScannerJig);
                        scaner.ReadData();
                        scaner.ReadData();
                    }
                    else
                        MessageBox.Show("Chưa có thông tin COM Scanner", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            catch (Exception error)
            {
                MessageBox.Show("Lỗi không kết nối được Scanner :\r\n" + error.ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }

        public void ReadCodeMaskBoard()
        {
            try
            {
                if (tb_Link5.BackColor == Color.Green)
                {
                retur:
                    scaner.ReadData();
                    code_MaskBoard = scaner.Data;

                    if (code_MaskBoard != null && code_MaskBoard != "" && code_MaskBoard != "ERROR")
                    {
                        ;
                    }
                    else
                    {
                        goto retur;
                    }

                    string[] listFileCode = Directory.GetFiles(@tb_Link6.Text, "*.txt");

                    if (listFileCode.Length > 0)
                    {
                        MessageBox.Show("Trong Folder " + tb_Link6.Text + " đã có 1 File Code PCM thông tin Sublead!!!!", "Lỗi thừa File Barcode PCM!!", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                    }
                    else
                    {
                        string data1 = "";
                        for (int i = 0; i < code_MaskBoard.Length; i++)
                        {
                            if (code_MaskBoard.Substring(i, 1) != "\r")
                                data1 += code_MaskBoard.Substring(i, 1);
                        }

                        code_MaskBoard = data1;

                        if (code_MaskBoard != null && code_MaskBoard != "" && code_MaskBoard != "ERROR")
                        {
                            string[] listFile = Directory.GetFiles(tb_Link5.Text, "*.txt");
                            if (listFile.Length > 0)
                            {
                                string[] _File = new string[listFile.Length];
                                for (int i = 0; i < listFile.Length; i++)
                                {
                                    string[] tmp = listFile[i].Split('\\');
                                    _File[i] = tmp[tmp.Length - 1];
                                }

                                checkCodeMaskboard = 0;
                                for (int i = 0; i < listFile.Length; i++)
                                {
                                    string[] data = _File[i].Split('-');
                                    if (data[0] == code_MaskBoard)
                                    {
                                        try
                                        {
                                            checkCodeMaskboard++;
                                            File.Move(listFile[i], tb_Link6.Text + @"\" + data[1]);
                                            break;
                                        }
                                        catch (Exception error)
                                        {
                                            MessageBox.Show("Lỗi khi đọc code Jig xong và tranfer File Log code\r\nKiểm tra lại 2 đường link số 5 và số 6\r\n " + error.ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                        }

                                    }
                                }

                                if (checkCodeMaskboard == 0)
                                {
                                    MessageBox.Show("Không tìm thấy File có Code Maskboard matching", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                }
                                else
                                {
                                    try
                                    {
                                    aa:
                                        string[] _list = Directory.GetFiles(@tb_Link6.Text, "*.txt");
                                        if (_list.Length > 0)
                                        {
                                            com_pc.sent_data("STRT");
                                            PLC.Writeplc("M205", 1);
                                            Thread.Sleep(300);
                                            PLC.Writeplc("M205", 0);
                                            count1 = 0;
                                        }
                                        else
                                        {
                                            goto aa;
                                        }
                                    }
                                    catch (Exception error)
                                    {
                                        MessageBox.Show("Lỗi khi kiểm tra folder lưu barcode function FU\r\nKiểm tra lại đường link số 6\r\n" + error.ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);

                                    }

                                }
                            }
                            else
                            {
                                DialogResult rsl = MessageBox.Show("Folder không có file Log!!!!!!!!"
                                                                          , "Thông báo", MessageBoxButtons.OKCancel, MessageBoxIcon.Error);
                            }
                        }
                        else
                        {
                            DialogResult rsl = MessageBox.Show("Không đọc được code Jig!!!!!!"
                                                                          , "Thông báo", MessageBoxButtons.OKCancel, MessageBoxIcon.Error);
                        }
                    }


                }
                else
                {
                    DialogResult rsl = MessageBox.Show("Không tìm thấy Folder chứa log File!!!!!!"
                                                                   , "Thông báo", MessageBoxButtons.OKCancel, MessageBoxIcon.Error);
                }
                tb_CodeMB.Text = "";
            }
            catch (Exception error)
            {
                MessageBox.Show("Lỗi đọc code maskboard :\r\n" + error.ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                tb_CodeMB.Text = "";
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            scaner.ReadData();
            tb_CodeMB.Text = scaner.Data;
        }

        private void button5_Click(object sender, EventArgs e)
        {
            ReadCodeMaskBoard();
        }

        private void tb_CodeMB_TextChanged(object sender, EventArgs e)
        {

        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            //btn_Disable.PerformClick();
        }

        public bool IsConnectInternet()
        {
            try
            {
                System.Net.IPHostEntry i = System.Net.Dns.GetHostEntry("107.107.147.177");
                return true;
            }
            catch
            {
                return false;            
            }
        }

        private void button5_Click_2(object sender, EventArgs e)
        {
            config.saveconfig_setup(tb_IPserver, tb_Link1, tb_Link2, tb_Link3, tb_Link4, tb_Link5, tb_Link6, tb_Ethernet, cbx_Scaner);
        }

        private void button6_Click(object sender, EventArgs e)
        {
            //string[] a = Directory.GetFiles(@tb_Link1.Text);
        }
    }
}
