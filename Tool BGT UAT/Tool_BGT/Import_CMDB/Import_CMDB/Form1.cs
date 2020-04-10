using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Xml.Linq;
using Excel_12 = Microsoft.Office.Interop.Excel;
using System.Data.OleDb;

namespace Import_CMDB
{
    public partial class frmImport : Form
    {
        #region định nghĩa các biến public
        public bool flag = true;
        public static ServiceReference.USD_WebServiceSoapClient ca;
        public static string ObjectResult, ObjectHandle;
        public static int SID;
        public static string DateNow;

        public static XDocument xDoc;


        public string[] NameColumnPOS = {"class","zcode","description","name","system_name",
                                            "zcategory_ref","zagency_ref","status","asset_count",
                                            "expiration_date","description","zstatus_pos","res_person",
                                            "res_tel","contact_person","contact_tel","zlocation_area_ref","delete_flag"};

        public string[] NameColumnTerminal = { "class", "name","system_name","description","zpo_date", "alarm_id",
                                                 "zhw_id","zterminal_number","acquire_date","install_date",
                                                 "zhoc","zpac","znote","zlocation_area_ref","zcode","status","delete_flag","expiration_date",
                                               "zsuspended_type","zdeploy_type"};

        public string[] NameColumnBGT = { "class", "name", "system_name", "zlocation_area_ref", "zcode",
                                            "alarm_id","zaddress_wanip_aon_ip", "zaddress_wanip_ip", 
                                            "zaddress_3g_wanip_ip","delete_flag" };

        public string[] NameColumnSim3G = { "class", "name", "system_name", "serial_number", 
                                              "zphone","active_date", "zlocation_area_ref","zisp_ref","family","delete_flag" };

        public string[] NameColumnLocation = { "class", "name", "system_name", 
                                                  "zcode","zDistrict", "zWard", "zStreet", 
                                                 "zHouse_Number","zlocation_area_ref","delete_flag"};
        public string[] NameColumnRouter = { "class", "name", "system_name", "serial_number", "zlocation_area_ref", "delete_flag" };

        public string[] NameColumnUPS = { "class", "name", "system_name", "serial_number", "zlocation_area_ref", "delete_flag" };

        public string[] NameColumnCPU = { "class", "name", "system_name", "serial_number", "zlocation_area_ref", "delete_flag" };

        public string[] NameColumnCDU = { "class", "name", "system_name", "serial_number", "zlocation_area_ref", "delete_flag" };

        public string[] NameColumnPrinter = { "class", "name", "system_name", "serial_number", "zlocation_area_ref", "delete_flag" };

        public string[] NameColumnScanner = { "class", "name", "system_name", "serial_number", "zlocation_area_ref", "delete_flag" };

        public string[] NameColumnKeyboard = { "class", "name", "system_name", "serial_number", "zlocation_area_ref", "delete_flag" };

        public string[] NameColumnTouchscreen = { "class", "name", "system_name", "serial_number", "zlocation_area_ref", "delete_flag" };

        public string[] NameColumnCheckwin = { "class", "name", "system_name", "serial_number", "ztype_ref", "zlocation_area_ref", "delete_flag" };

        public string[] NameColumnKho = { "class", "name", "system_name", "zlocation_area_ref", "delete_flag" };


        public string[] NameColumnRelation = { "pos", "bgt", "location", "terminal", "sim3g","checkwin",
                                                 "router","ups","cpu","cdu","printer","scanner","keyboard",
                                                 "touchscreen","zlocation_area_ref" };

        public string[] NameColumnRela_POS = { "id", "class", "connects to" };

        public string[] NameColumnRela_POS_Add = { "id", "class", "is location for" };

        public string[] NameColumnRela_POS_Termi = { "id", "class", "contains" };

        public string[] NameColumnRela_Terminal_Other = { "id", "class", "contains", "contains", "contains", "contains",
                                                            "contains", "contains", "contains", "contains", "contains", "contains" };

        public string[] NameColumnRela_kho_other = { "id", "class", "is location for", "is location for", "is location for", "is location for", "is location for",
                                                       "is location for", "is location for", "is location for", "is location for", "is location for", "is location for"};

        public string[] NameSheetTemplate = { "FullCI", "BGT", "Sim3G" };

        //is location for
        #endregion

        public frmImport()
        {
            InitializeComponent();
        }

        /// <summary>
        ///  Ghi log Chuong trinh vào file log.txt trong thư mục bin
        /// </summary>
        /// <param name="value"></param>
        public void WriteOLog(string value)
        {
            using (StreamWriter write = new StreamWriter(txtFolderLog.Text + "\\log.txt", true))
            {
                write.WriteLine(value);
            }
        }

        private void btnDangNhap_Click(object sender, EventArgs e)
        {
            DateNow = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss tt");
            WriteOLog("\t ----------------- Truy cập phần mềm - Date Time: " + DateNow + "-----------------");

            if (checkEmpty() == 0) return;
            if (Check_Config())
            {
                MessageBox.Show("Đăng nhập CA Service Desk thành công!", "Message");
                WriteOLog("Đăng nhập CA Service Desk thành công!");

                tabControl1.TabPages.Add(tabPage2);

            }
            else
            {
                MessageBox.Show("Đăng nhập CA Service Desk không thành công!", "Message");
                WriteOLog("Đăng nhập CA Service Desk không thành công!");
            }
        }

        private int checkEmpty()
        {
            if (string.IsNullOrEmpty(txtUser.Text) || string.IsNullOrEmpty(txtPass.Text) || string.IsNullOrEmpty(txtFolderLog.Text) || string.IsNullOrEmpty(txtFolderGrloader.Text))
            {
                MessageBox.Show("Yêu cầu nhập các trường tin bắt buộc.", "Message");
                WriteOLog("Yêu cầu nhập các trường tin bắt buộc.");

                return 0;
            }
            return 1;
        }

        /// <summary>
        /// Check đăng nhập CA
        /// 
        /// </summary>
        /// <returns></returns>
        public bool Check_Config()
        {
            string User = txtUser.Text.Trim();
            string Pass = txtPass.Text.Trim();
            #region Kết nối CA
            try
            {
                ca = new ServiceReference.USD_WebServiceSoapClient();
                xDoc = new XDocument();
                SID = ca.login(User, Pass);
                WriteOLog("Đăng nhập CA thành công");
            }
            catch
            {
                WriteOLog("Đăng nhập CA không thành công");
                return false;
            }
            return true;
            #endregion
        }

        /// <summary>
        /// Sự kiện chọn thư mục cài đặt GRloader
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnChonGrloader_Click(object sender, EventArgs e)
        {
            DialogResult result = folderBrowserDialog1.ShowDialog();
            if (result == DialogResult.OK)
            {
                string txt = folderBrowserDialog1.SelectedPath.ToString();
                txtFolderGrloader.Text = txt;
            }
        }

        private void btnChonLog_Click(object sender, EventArgs e)
        {
            DialogResult result = folderBrowserDialog1.ShowDialog();
            if (result == DialogResult.OK)
            {
                string txt = folderBrowserDialog1.SelectedPath.ToString();
                txtFolderLog.Text = txt;
            }
        }

        private void frmImport_Load(object sender, EventArgs e)
        {
            

            tabControl1.TabPages.Remove(tabPage2);

        }

        private void btnThoat_Click(object sender, EventArgs e)
        {
            DateNow = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss tt");
            WriteOLog("\t ----------------- Đăng xuất hệ thống - Date Time: " + DateNow + "-----------------");
            this.Close();
        }


        #region Tạo file import dữ liệu từng loại
        public void CreatePOS(string strNameFileNew, string strNameSheet, DataTable dbFull)
        {
            Microsoft.Office.Interop.Excel.Application excel;
            Microsoft.Office.Interop.Excel.Workbook worKbooK;
            Microsoft.Office.Interop.Excel.Worksheet worKsheeT;
            Microsoft.Office.Interop.Excel.Range celLrangE;

            try
            {
                excel = new Microsoft.Office.Interop.Excel.Application();
                excel.Interactive = false;
                excel.Visible = false;
                excel.DisplayAlerts = false;
                worKbooK = excel.Workbooks.Add(Type.Missing);


                worKsheeT = (Microsoft.Office.Interop.Excel.Worksheet)worKbooK.ActiveSheet;
                worKsheeT.Name = strNameSheet;

                for (int i = 1; i <= NameColumnPOS.Length; i++)
                {
                    worKsheeT.Cells[1, i] = NameColumnPOS[i - 1];
                }


                if (dbFull != null && dbFull.Rows.Count > 0)
                {
                    for (int i = 1; i <= dbFull.Rows.Count; i++)
                    {
                        string posid = dbFull.Rows[i - 1][0].ToString();

                        if (!string.IsNullOrEmpty(posid))
                        {
                            #region tao sheet POS

                            // Class
                            worKsheeT.Cells[i + 1, 1] = "POS";

                            // POS ID = zcode
                            worKsheeT.Cells[i + 1, 2] = "'" + posid.ToString();

                            //Description = note pos
                            string description = dbFull.Rows[i - 1][42].ToString();
                            worKsheeT.Cells[i + 1, 3] = (!string.IsNullOrEmpty(description)) ? description : "EMPTY";  
                            //POS name 
                            worKsheeT.Cells[i + 1, 4] = "POS-" + posid;

                            // System name 
                            worKsheeT.Cells[i + 1, 5] = "POS-" + posid;

                            //zcategory_ref
                            string postype = dbFull.Rows[i - 1][3].ToString();
                            worKsheeT.Cells[i + 1, 6] =  (!string.IsNullOrEmpty(postype)) ? postype : "EMPTY";

                            //zagency_ref
                            string postagence = dbFull.Rows[i - 1][4].ToString();
                            worKsheeT.Cells[i + 1, 7] = (!string.IsNullOrEmpty(postagence)) ? postagence : "EMPTY";

                            //status
                            //string poststatus = dbFull.Rows[i - 1][2].ToString();
                           // worKsheeT.Cells[i + 1, 8] = poststatus;

                            //zstatus_pos
                            string zstatus_pos = dbFull.Rows[i - 1][38].ToString();
                            worKsheeT.Cells[i + 1, 8] = (!string.IsNullOrEmpty(zstatus_pos)) ? zstatus_pos : "EMPTY";

                            //"res_person"
                            string res_person = dbFull.Rows[i - 1][11].ToString();
                            worKsheeT.Cells[i + 1, 13] = (!string.IsNullOrEmpty(res_person)) ? res_person : "EMPTY"; 

                            //,"res_tel",
                            string res_tel = dbFull.Rows[i - 1][12].ToString();
                            worKsheeT.Cells[i + 1, 14] = (!string.IsNullOrEmpty(res_tel)) ? res_tel : "EMPTY";  

                            //"contact_person",
                            string contact_person = dbFull.Rows[i-1][13].ToString();
                            worKsheeT.Cells[i + 1, 15] = (!string.IsNullOrEmpty(contact_person)) ? contact_person : "EMPTY";  
                            //"contact_tel"
                            string contact_tel = dbFull.Rows[i - 1][14].ToString();
                            worKsheeT.Cells[i + 1, 16] = (!string.IsNullOrEmpty(contact_tel)) ? contact_tel : "EMPTY";  
                            //,"zlocation_area_ref"
                            string zlocation_area_ref = dbFull.Rows[i - 1][10].ToString();
                            worKsheeT.Cells[i + 1, 17] = (!string.IsNullOrEmpty(zlocation_area_ref)) ? zlocation_area_ref : "EMPTY";
                            //active
                            worKsheeT.Cells[i + 1, 18] = "0";
                            #endregion
                        }

                    }
                }

                worKbooK.SaveAs(strNameFileNew); ;
                worKbooK.Close();
                excel.Quit();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
                throw;
            }


        }

        public void CreateTerminal(string strNameFileNew, string strNameSheet, DataTable dbFull)
        {
            Microsoft.Office.Interop.Excel.Application excel;
            Microsoft.Office.Interop.Excel.Workbook worKbooK;
            Microsoft.Office.Interop.Excel.Worksheet worKsheeT;
            Microsoft.Office.Interop.Excel.Range celLrangE;

            try
            {
                excel = new Microsoft.Office.Interop.Excel.Application();
                excel.Visible = false;
                excel.DisplayAlerts = false;
                worKbooK = excel.Workbooks.Add(Type.Missing);


                worKsheeT = (Microsoft.Office.Interop.Excel.Worksheet)worKbooK.ActiveSheet;
                worKsheeT.Name = strNameSheet;

                for (int i = 1; i <= NameColumnTerminal.Length; i++)
                {
                    worKsheeT.Cells[1, i] = NameColumnTerminal[i - 1];
                }


                if (dbFull != null && dbFull.Rows.Count > 0)
                {
                    for (int i = 1; i <= dbFull.Rows.Count; i++)
                    {
                        string zhw_id = dbFull.Rows[i - 1][19].ToString();
                        string name = dbFull.Rows[i - 1][20].ToString();
                        string status = dbFull.Rows[i - 1][2].ToString();

                        if (!string.IsNullOrEmpty(name))
                        {
                            #region tao sheet Terminal
                            // Class
                            worKsheeT.Cells[i + 1, 1] = "Terminal";

                            //name

                            worKsheeT.Cells[i + 1, 2] = name;
                            // system_name 
                            worKsheeT.Cells[i + 1, 3] = name;
                            // description (doi tu znote)
                            string description = dbFull.Rows[i - 1][41].ToString();
                            worKsheeT.Cells[i + 1, 4] = (!string.IsNullOrEmpty(description)) ? ("'" + description) : "EMPTY"; 
                            //zpo_date
                            string zpo_date = dbFull.Rows[i - 1][15].ToString();
                            worKsheeT.Cells[i + 1, 5] = zpo_date;
                            //alarm_id
                            string alarm_id = dbFull.Rows[i - 1][18].ToString();
                            worKsheeT.Cells[i + 1, 6] = alarm_id;
                            //zhw_id

                            worKsheeT.Cells[i + 1, 7] = (!string.IsNullOrEmpty(zhw_id)) ? ("'" + zhw_id) : "EMPTY"; 
                            //zterminal_number
                            string zterminal_number = dbFull.Rows[i - 1][21].ToString();
                            worKsheeT.Cells[i + 1, 8] = (!string.IsNullOrEmpty(zterminal_number)) ? zterminal_number : "EMPTY";
                            //acquire_date
                            string acquire_date = dbFull.Rows[i - 1][22].ToString();
                            try
                            {
                                acquire_date = (!string.IsNullOrEmpty(acquire_date)) ? Convert.ToDateTime(acquire_date).ToShortDateString() : "EMPTY";
                            }
                            catch (Exception)
                            {
                                acquire_date = "EMPTY";
                                throw;
                            }
                            worKsheeT.Cells[i + 1, 9] = acquire_date;
                            //install_date
                            string install_date = dbFull.Rows[i - 1][23].ToString();
                            try
                            {
                                install_date = (!string.IsNullOrEmpty(install_date)) ? Convert.ToDateTime(install_date).ToShortDateString() : "EMPTY";
                            }
                            catch (Exception)
                            {
                                install_date = "EMPTY";
                                throw;
                            }

                            worKsheeT.Cells[i + 1, 10] = install_date;
                            //zhoc
                            string zhoc = dbFull.Rows[i - 1][39].ToString();
                            worKsheeT.Cells[i + 1, 11] = zhoc.Length > 0 ? "1" : "0";
                            //zpac
                            string zpac = dbFull.Rows[i - 1][40].ToString();
                            worKsheeT.Cells[i + 1, 12] = zpac.Length > 0 ? "1" : "0";
                            //znote
                            //string znote = dbFull.Rows[i - 1][41].ToString();
                            //worKsheeT.Cells[i + 1, 13] = znote;
                            //,"zlocation_area_ref"
                            string zlocation_area_ref = dbFull.Rows[i - 1][10].ToString();
                            worKsheeT.Cells[i + 1, 14] = (!string.IsNullOrEmpty(zlocation_area_ref)) ? zlocation_area_ref : "EMPTY";  
                            //zcode
                            worKsheeT.Cells[i + 1, 15] = (!string.IsNullOrEmpty(name)) ? name : "EMPTY";  
                            //status                           
                            worKsheeT.Cells[i + 1, 16] = (!string.IsNullOrEmpty(status)) ? status : "EMPTY"; ;
                            //active
                            worKsheeT.Cells[i + 1, 17] = "0";
                            // Supsended Date
                            string Supsended_Date = dbFull.Rows[i - 1][44].ToString();
                            try
                            {
                                Supsended_Date = (!string.IsNullOrEmpty(Supsended_Date)) ? Convert.ToDateTime(Supsended_Date).ToShortDateString() : "EMPTY";
                            }
                            catch (Exception)
                            {
                                Supsended_Date = "EMPTY";
                                throw;
                            }
                            worKsheeT.Cells[i + 1, 18] = Supsended_Date;
                            // Suspended type
                            string Suspended_type = dbFull.Rows[i - 1][45].ToString();
                            worKsheeT.Cells[i + 1, 19] = (!string.IsNullOrEmpty(Suspended_type)) ? Suspended_type : "EMPTY";  

                            //Deployment Type
                            string Deployment_Type = dbFull.Rows[i - 1][46].ToString();
                            worKsheeT.Cells[i + 1, 20] = (!string.IsNullOrEmpty(Deployment_Type)) ? Deployment_Type : "EMPTY";  
                            #endregion
                        }
                    }
                }

                worKbooK.SaveAs(strNameFileNew); ;
                worKbooK.Close();
                excel.Quit();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
                throw;
            }

        }

        public void CreateBGTCode(string strNameFileNew, string strNameSheet, DataTable dbBGT)
        {
            Microsoft.Office.Interop.Excel.Application excel;
            Microsoft.Office.Interop.Excel.Workbook worKbooK;
            Microsoft.Office.Interop.Excel.Worksheet worKsheeT;
            Microsoft.Office.Interop.Excel.Range celLrangE;

            try
            {
                excel = new Microsoft.Office.Interop.Excel.Application();
                excel.Visible = false;
                excel.DisplayAlerts = false;
                worKbooK = excel.Workbooks.Add(Type.Missing);


                worKsheeT = (Microsoft.Office.Interop.Excel.Worksheet)worKbooK.ActiveSheet;
                worKsheeT.Name = strNameSheet;

                for (int i = 1; i <= NameColumnBGT.Length; i++)
                {
                    worKsheeT.Cells[1, i] = NameColumnBGT[i - 1];
                }


                if (dbBGT != null && dbBGT.Rows.Count > 0)
                {
                    for (int i = 1; i <= dbBGT.Rows.Count; i++)
                    {
                        string name = dbBGT.Rows[i - 1][2].ToString();
                        if (!string.IsNullOrEmpty(name))
                        {
                            #region tao sheet Terminal
                            // Class
                            worKsheeT.Cells[i + 1, 1] = "BGT Code";
                            // name 

                            worKsheeT.Cells[i + 1, 2] = name;
                            //system_name 
                            worKsheeT.Cells[i + 1, 3] = name;
                            //zlocation_area_ref
                            string location = dbBGT.Rows[i - 1][1].ToString();
                            worKsheeT.Cells[i + 1, 4] = (!string.IsNullOrEmpty(location)) ? location : "EMPTY"; 
                            //bgt code
                            worKsheeT.Cells[i + 1, 5] = name;
                            //alarm_id
                            string alarm_id = dbBGT.Rows[i - 1][3].ToString();
                            worKsheeT.Cells[i + 1, 6] = (!string.IsNullOrEmpty(alarm_id)) ? alarm_id : "EMPTY";
                            //zaddress_wanip_aon_ip
                            string zaddress_wanip_aon_ip = dbBGT.Rows[i - 1][5].ToString();
                            worKsheeT.Cells[i + 1, 7] = (!string.IsNullOrEmpty(zaddress_wanip_aon_ip)) ? zaddress_wanip_aon_ip : "EMPTY"; 
                            //zaddress_wanip_ip
                            string zaddress_wanip_ip = dbBGT.Rows[i - 1][6].ToString();
                            worKsheeT.Cells[i + 1, 8] = (!string.IsNullOrEmpty(zaddress_wanip_ip)) ? zaddress_wanip_ip : "EMPTY"; 
                            //zaddress_3g_wanip_ip
                            string zaddress_3g_wanip_ip = dbBGT.Rows[i - 1][7].ToString();
                            worKsheeT.Cells[i + 1, 9] = (!string.IsNullOrEmpty(zaddress_3g_wanip_ip)) ? zaddress_3g_wanip_ip : "EMPTY";  
                            //active
                            worKsheeT.Cells[i + 1, 10] = "0";
                            #endregion
                        }
                    }
                }

                worKbooK.SaveAs(strNameFileNew); ;
                worKbooK.Close();
                excel.Quit();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
                throw;
            }


        }

        public void CreateSim3G(string strNameFileNew, string strNameSheet, DataTable dbSim3G)
        {
            Microsoft.Office.Interop.Excel.Application excel;
            Microsoft.Office.Interop.Excel.Workbook worKbooK;
            Microsoft.Office.Interop.Excel.Worksheet worKsheeT;
            Microsoft.Office.Interop.Excel.Range celLrangE;

            try
            {
                excel = new Microsoft.Office.Interop.Excel.Application();
                excel.Visible = false;
                excel.DisplayAlerts = false;
                worKbooK = excel.Workbooks.Add(Type.Missing);


                worKsheeT = (Microsoft.Office.Interop.Excel.Worksheet)worKbooK.ActiveSheet;
                worKsheeT.Name = strNameSheet;

                for (int i = 1; i <= NameColumnSim3G.Length; i++)
                {
                    worKsheeT.Cells[1, i] = NameColumnSim3G[i - 1];
                }


                if (dbSim3G != null && dbSim3G.Rows.Count > 0)
                {
                    for (int i = 1; i <= dbSim3G.Rows.Count; i++)
                    {
                        string name = dbSim3G.Rows[i - 1][0].ToString().Trim();

                        if (!string.IsNullOrEmpty(name))
                        {
                            // 
                             //                  "zlocation_area_ref","zisp_ref" 

                            #region tao sheet Terminal
                            // Class
                            worKsheeT.Cells[i + 1, 1] = "SIM 3G".Trim();
                            // name 
                            worKsheeT.Cells[i + 1, 2] = "SIM 3G-" + name;
                            //system_name
                            worKsheeT.Cells[i + 1, 3] = "SIM 3G-" + name;
                            //serial_number
                            worKsheeT.Cells[i + 1, 4] = "'" + name;
                            //zphone
                            string zphone = dbSim3G.Rows[i - 1][1].ToString();
                            worKsheeT.Cells[i + 1, 5] = (!string.IsNullOrEmpty(zphone)) ? zphone : "EMPTY"; 
                            //active_date
                            string active_date = dbSim3G.Rows[i - 1][2].ToString();
                            try
                            {
                                active_date = (!string.IsNullOrEmpty(active_date)) ? Convert.ToDateTime(active_date).ToShortDateString() : "EMPTY";
                            }
                            catch (Exception)
                            {
                                active_date = "EMPTY";
                                throw;
                            }
                            worKsheeT.Cells[i + 1, 6] = active_date;                           
                            //,"zlocation_area_ref"
                            string zlocation_area_ref = dbSim3G.Rows[i - 1][3].ToString();
                            worKsheeT.Cells[i + 1, 7] = (!string.IsNullOrEmpty(zlocation_area_ref)) ? zlocation_area_ref : "EMPTY";  
                            //isp
                            string isp = dbSim3G.Rows[i - 1][5].ToString();
                            worKsheeT.Cells[i + 1, 8] = (!string.IsNullOrEmpty(isp)) ? isp : "EMPTY";  
                            //family
                            worKsheeT.Cells[i + 1, 9] = "Network";
                            //active
                            worKsheeT.Cells[i + 1, 10] = "0";
                            #endregion
                        }
                    }
                }

                worKbooK.SaveAs(strNameFileNew); ;
                worKbooK.Close();
                excel.Quit();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
                throw;
            }

        }

        public void CreateLocation(string strNameFileNew, string strNameSheet, DataTable dbLocation)
        {
            Microsoft.Office.Interop.Excel.Application excel;
            Microsoft.Office.Interop.Excel.Workbook worKbooK;
            Microsoft.Office.Interop.Excel.Worksheet worKsheeT;
            Microsoft.Office.Interop.Excel.Range celLrangE;

            try
            {
                excel = new Microsoft.Office.Interop.Excel.Application();
                excel.Visible = false;
                excel.DisplayAlerts = false;
                worKbooK = excel.Workbooks.Add(Type.Missing);


                worKsheeT = (Microsoft.Office.Interop.Excel.Worksheet)worKbooK.ActiveSheet;
                worKsheeT.Name = strNameSheet;

                for (int i = 1; i <= NameColumnLocation.Length; i++)
                {
                    worKsheeT.Cells[1, i] = NameColumnLocation[i - 1];
                }


                if (dbLocation != null && dbLocation.Rows.Count > 0)
                {
                    for (int i = 1; i <= dbLocation.Rows.Count; i++)
                    {
                        string name = dbLocation.Rows[i - 1][5].ToString();

                        // kiểm tra address này đã có trên hệ thống hay chưa 
                        List<string> lstAddId = GetDataID(name, "name", "nr");
                        // Nếu trường name add trong excel có dữ liệu, và dữ liệu này không tồn tại trên CA thì đẩy vào file import
                        if (!string.IsNullOrEmpty(name) && (lstAddId == null || lstAddId.Count == 0))
                        {
                            #region tao sheet Address
                            // Class
                            worKsheeT.Cells[i + 1, 1] = "Address";
                            // name 

                            worKsheeT.Cells[i + 1, 2] = name;
                            //system_name
                            worKsheeT.Cells[i + 1, 3] = name;
                            //zcode
                            //oRange = (Excel_12.Range)oSheet.Cells[i + 1, 4];
                            //oRange.Value2 = name;
                            //District
                            string District = dbLocation.Rows[i - 1][9].ToString();
                            worKsheeT.Cells[i + 1, 5] = (!string.IsNullOrEmpty(District)) ? District : "EMPTY"; ;
                            //Ward*
                            string Ward = dbLocation.Rows[i - 1][8].ToString();
                            worKsheeT.Cells[i + 1, 6] = (!string.IsNullOrEmpty(Ward)) ? Ward : "EMPTY"; 
                            //Street
                            string Street = dbLocation.Rows[i - 1][7].ToString();
                            worKsheeT.Cells[i + 1, 7] = (!string.IsNullOrEmpty(Street)) ? Street : "EMPTY";
                            //zHouse_Number
                            string HouseNumber = dbLocation.Rows[i - 1][6].ToString();
                            worKsheeT.Cells[i + 1, 8] = (!string.IsNullOrEmpty(HouseNumber)) ? HouseNumber : "EMPTY";
                            //,"zlocation_area_ref"
                            string zlocation_area_ref = dbLocation.Rows[i - 1][10].ToString();
                            worKsheeT.Cells[i + 1, 9] = (!string.IsNullOrEmpty(zlocation_area_ref)) ? zlocation_area_ref : "EMPTY"; 
                            //active
                            worKsheeT.Cells[i + 1, 10] = "0";
                            #endregion
                        }
                    }
                }

                worKbooK.SaveAs(strNameFileNew); ;
                worKbooK.Close();
                excel.Quit();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
                throw;
            }
        }

        public void CreateRouter(string strNameFileNew, string strNameSheet, DataTable bdFull)
        {
            Microsoft.Office.Interop.Excel.Application excel;
            Microsoft.Office.Interop.Excel.Workbook worKbooK;
            Microsoft.Office.Interop.Excel.Worksheet worKsheeT;
            Microsoft.Office.Interop.Excel.Range celLrangE;

            try
            {
                excel = new Microsoft.Office.Interop.Excel.Application();
                excel.Visible = false;
                excel.DisplayAlerts = false;
                worKbooK = excel.Workbooks.Add(Type.Missing);


                worKsheeT = (Microsoft.Office.Interop.Excel.Worksheet)worKbooK.ActiveSheet;
                worKsheeT.Name = strNameSheet;

                for (int i = 1; i <= NameColumnRouter.Length; i++)
                {
                    worKsheeT.Cells[1, i] = NameColumnRouter[i - 1];
                }

                if (bdFull != null && bdFull.Rows.Count > 0)
                {
                    for (int i = 1; i <= bdFull.Rows.Count; i++)
                    {
                        string name = bdFull.Rows[i - 1][28].ToString();
                        string terminalid = bdFull.Rows[i - 1][20].ToString();

                        if (!string.IsNullOrEmpty(terminalid) && !string.IsNullOrEmpty(name))
                        {
                            #region tao sheet Terminal
                            // Class
                            worKsheeT.Cells[i + 1, 1] = "Router";
                            // name 

                            worKsheeT.Cells[i + 1, 2] = "Router-" + name;
                            //system_name
                            //worKsheeT.Cells[i + 1, 3] = "Router-" + name;
                            //serial_number
                            worKsheeT.Cells[i + 1, 4] = name;
                            //,"zlocation_area_ref"
                            string zlocation_area_ref = bdFull.Rows[i - 1][10].ToString();
                            worKsheeT.Cells[i + 1, 5] = (!string.IsNullOrEmpty(zlocation_area_ref)) ? zlocation_area_ref : "EMPTY";
                            // active
                            worKsheeT.Cells[i + 1, 6] = "0";
                            #endregion
                        }
                    }
                }
                worKbooK.SaveAs(strNameFileNew); ;
                worKbooK.Close();
                excel.Quit();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
                throw;
            }

        }

        public void CreateUPS(string strNameFileNew, string strNameSheet, DataTable bdFull)
        {


            Microsoft.Office.Interop.Excel.Application excel;
            Microsoft.Office.Interop.Excel.Workbook worKbooK;
            Microsoft.Office.Interop.Excel.Worksheet worKsheeT;
            Microsoft.Office.Interop.Excel.Range celLrangE;

            try
            {
                excel = new Microsoft.Office.Interop.Excel.Application();
                excel.Visible = false;
                excel.DisplayAlerts = false;
                worKbooK = excel.Workbooks.Add(Type.Missing);


                worKsheeT = (Microsoft.Office.Interop.Excel.Worksheet)worKbooK.ActiveSheet;
                worKsheeT.Name = strNameSheet;

                for (int i = 1; i <= NameColumnUPS.Length; i++)
                {
                    worKsheeT.Cells[1, i] = NameColumnUPS[i - 1];
                }


                if (bdFull != null && bdFull.Rows.Count > 0)
                {
                    for (int i = 1; i <= bdFull.Rows.Count; i++)
                    {
                        string name = bdFull.Rows[i - 1][29].ToString();
                        string terminalid = bdFull.Rows[i - 1][20].ToString();

                        if (!string.IsNullOrEmpty(terminalid) && !string.IsNullOrEmpty(name))
                        {
                            #region tao sheet Terminal
                            // Class
                            worKsheeT.Cells[i + 1, 1] = "UPS";
                            // name                     
                            worKsheeT.Cells[i + 1, 2] = "UPS-" + name;
                            //system_name
                            //worKsheeT.Cells[i + 1, 3] = "UPS-" + name;
                            //serial_number
                            worKsheeT.Cells[i + 1, 4] = "'" + name;
                            //,"zlocation_area_ref"
                            string zlocation_area_ref = bdFull.Rows[i - 1][10].ToString();
                            worKsheeT.Cells[i + 1, 5] = zlocation_area_ref;
                            //active
                            worKsheeT.Cells[i + 1, 6] = "0";
                            #endregion
                        }
                    }
                }
                worKbooK.SaveAs(strNameFileNew); ;
                worKbooK.Close();
                excel.Quit();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
                throw;
            }

        }

        public void CreateCPU(string strNameFileNew, string strNameSheet, DataTable bdFull)
        {

            Microsoft.Office.Interop.Excel.Application excel;
            Microsoft.Office.Interop.Excel.Workbook worKbooK;
            Microsoft.Office.Interop.Excel.Worksheet worKsheeT;
            Microsoft.Office.Interop.Excel.Range celLrangE;

            try
            {
                excel = new Microsoft.Office.Interop.Excel.Application();
                excel.Visible = false;
                excel.DisplayAlerts = false;
                worKbooK = excel.Workbooks.Add(Type.Missing);


                worKsheeT = (Microsoft.Office.Interop.Excel.Worksheet)worKbooK.ActiveSheet;
                worKsheeT.Name = strNameSheet;

                for (int i = 1; i <= NameColumnCPU.Length; i++)
                {
                    worKsheeT.Cells[1, i] = NameColumnCPU[i - 1];
                }


                if (bdFull != null && bdFull.Rows.Count > 0)
                {
                    for (int i = 1; i <= bdFull.Rows.Count; i++)
                    {
                        string name = bdFull.Rows[i - 1][30].ToString();
                        string terminalid = bdFull.Rows[i - 1][20].ToString();

                        if (!string.IsNullOrEmpty(terminalid) && !string.IsNullOrEmpty(name))
                        {
                            #region tao sheet Terminal
                            // Class
                            worKsheeT.Cells[i + 1, 1] = "CPU";
                            // name 
                            worKsheeT.Cells[i + 1, 2] = "CPU-" + name;
                            //system_name
                            //worKsheeT.Cells[i + 1, 3] = "CPU-" + name;
                            //serial_number
                            worKsheeT.Cells[i + 1, 4] = name;
                            //,"zlocation_area_ref"
                            string zlocation_area_ref = bdFull.Rows[i - 1][10].ToString();
                            worKsheeT.Cells[i + 1, 5] = zlocation_area_ref;
                            //active
                            worKsheeT.Cells[i + 1, 6] = "0";
                            #endregion
                        }
                    }
                }

                worKbooK.SaveAs(strNameFileNew); ;
                worKbooK.Close();
                excel.Quit();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
                throw;
            }
        }

        public void CreateCDU(string strNameFileNew, string strNameSheet, DataTable bdFull)
        {
            Microsoft.Office.Interop.Excel.Application excel;
            Microsoft.Office.Interop.Excel.Workbook worKbooK;
            Microsoft.Office.Interop.Excel.Worksheet worKsheeT;
            Microsoft.Office.Interop.Excel.Range celLrangE;

            try
            {
                excel = new Microsoft.Office.Interop.Excel.Application();
                excel.Visible = false;
                excel.DisplayAlerts = false;
                worKbooK = excel.Workbooks.Add(Type.Missing);


                worKsheeT = (Microsoft.Office.Interop.Excel.Worksheet)worKbooK.ActiveSheet;
                worKsheeT.Name = strNameSheet;

                for (int i = 1; i <= NameColumnCDU.Length; i++)
                {
                    worKsheeT.Cells[1, i] = NameColumnCDU[i - 1];
                }


                if (bdFull != null && bdFull.Rows.Count > 0)
                {
                    for (int i = 1; i <= bdFull.Rows.Count; i++)
                    {
                        string name = bdFull.Rows[i - 1][31].ToString();
                        string terminalid = bdFull.Rows[i - 1][20].ToString();

                        if (!string.IsNullOrEmpty(terminalid) && !string.IsNullOrEmpty(name))
                        {
                            #region tao sheet Terminal
                            // Class
                            worKsheeT.Cells[i + 1, 1] = "CDU";
                            // name                     
                            worKsheeT.Cells[i + 1, 2] = "CDU-" + name;
                            //system_name
                            //worKsheeT.Cells[i + 1, 3] = "CDU-" + name;
                            //serial_number
                            worKsheeT.Cells[i + 1, 4] = "'" + name;
                            //,"zlocation_area_ref"
                            string zlocation_area_ref = bdFull.Rows[i - 1][10].ToString();
                            worKsheeT.Cells[i + 1, 5] = zlocation_area_ref;
                            //active
                            worKsheeT.Cells[i + 1, 6] = "0";
                            #endregion
                        }
                    }
                }

                worKbooK.SaveAs(strNameFileNew); ;
                worKbooK.Close();
                excel.Quit();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
                throw;
            }


        }

        public void CreatePrinter(string strNameFileNew, string strNameSheet, DataTable bdFull)
        {
            Microsoft.Office.Interop.Excel.Application excel;
            Microsoft.Office.Interop.Excel.Workbook worKbooK;
            Microsoft.Office.Interop.Excel.Worksheet worKsheeT;
            Microsoft.Office.Interop.Excel.Range celLrangE;

            try
            {
                excel = new Microsoft.Office.Interop.Excel.Application();
                excel.Visible = false;
                excel.DisplayAlerts = false;
                worKbooK = excel.Workbooks.Add(Type.Missing);


                worKsheeT = (Microsoft.Office.Interop.Excel.Worksheet)worKbooK.ActiveSheet;
                worKsheeT.Name = strNameSheet;

                for (int i = 1; i <= NameColumnPrinter.Length; i++)
                {
                    worKsheeT.Cells[1, i] = NameColumnPrinter[i - 1];
                }

        #endregion

                if (bdFull != null && bdFull.Rows.Count > 0)
                {
                    for (int i = 1; i <= bdFull.Rows.Count; i++)
                    {
                        string name = bdFull.Rows[i - 1][32].ToString();
                        string terminalid = bdFull.Rows[i - 1][20].ToString();

                        if (!string.IsNullOrEmpty(terminalid) && !string.IsNullOrEmpty(name))
                        {
                            #region tao sheet Terminal
                            // Class
                            worKsheeT.Cells[i + 1, 1] = "Printer";
                            // name                     
                            worKsheeT.Cells[i + 1, 2] = "Printer-" + name;
                            //system_name
                            //worKsheeT.Cells[i + 1, 3] = "Printer-" + name;
                            //serial_number
                            worKsheeT.Cells[i + 1, 4] = "'" + name;
                            //,"zlocation_area_ref"
                            string zlocation_area_ref = bdFull.Rows[i - 1][10].ToString();
                            worKsheeT.Cells[i + 1, 5] = zlocation_area_ref;
                            //active
                            worKsheeT.Cells[i + 1, 6] = "0";
                            #endregion
                        }
                    }
                }
                worKbooK.SaveAs(strNameFileNew); ;
                worKbooK.Close();
                excel.Quit();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
                throw;
            }

        }

        public void CreateScanner(string strNameFileNew, string strNameSheet, DataTable bdFull)
        {

            Microsoft.Office.Interop.Excel.Application excel;
            Microsoft.Office.Interop.Excel.Workbook worKbooK;
            Microsoft.Office.Interop.Excel.Worksheet worKsheeT;
            Microsoft.Office.Interop.Excel.Range celLrangE;

            try
            {
                excel = new Microsoft.Office.Interop.Excel.Application();
                excel.Visible = false;
                excel.DisplayAlerts = false;
                worKbooK = excel.Workbooks.Add(Type.Missing);


                worKsheeT = (Microsoft.Office.Interop.Excel.Worksheet)worKbooK.ActiveSheet;
                worKsheeT.Name = strNameSheet;

                for (int i = 1; i <= NameColumnScanner.Length; i++)
                {
                    worKsheeT.Cells[1, i] = NameColumnScanner[i - 1];
                }

                if (bdFull != null && bdFull.Rows.Count > 0)
                {
                    for (int i = 1; i <= bdFull.Rows.Count; i++)
                    {
                        string name = bdFull.Rows[i - 1][33].ToString();
                        string terminalid = bdFull.Rows[i - 1][20].ToString();

                        if (!string.IsNullOrEmpty(terminalid) && !string.IsNullOrEmpty(name))
                        {
                            #region tao sheet Terminal
                            // Class
                            worKsheeT.Cells[i + 1, 1] = "Scanner";
                            // name                    
                            worKsheeT.Cells[i + 1, 2] = "Scanner-" + name;
                            //system_name
                            //worKsheeT.Cells[i + 1, 3] = "Scanner-" + name;
                            //serial_number
                            worKsheeT.Cells[i + 1, 4] = name;
                            //,"zlocation_area_ref"
                            string zlocation_area_ref = bdFull.Rows[i - 1][10].ToString();
                            worKsheeT.Cells[i + 1, 5] = zlocation_area_ref;
                            //active
                            worKsheeT.Cells[i + 1, 6] = "0";
                            #endregion
                        }
                    }
                }
                worKbooK.SaveAs(strNameFileNew); ;
                worKbooK.Close();
                excel.Quit();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
                throw;
            }
        }

        public void CreateKeyboard(string strNameFileNew, string strNameSheet, DataTable bdFull)
        {
            Microsoft.Office.Interop.Excel.Application excel;
            Microsoft.Office.Interop.Excel.Workbook worKbooK;
            Microsoft.Office.Interop.Excel.Worksheet worKsheeT;
            Microsoft.Office.Interop.Excel.Range celLrangE;

            try
            {
                excel = new Microsoft.Office.Interop.Excel.Application();
                excel.Visible = false;
                excel.DisplayAlerts = false;
                worKbooK = excel.Workbooks.Add(Type.Missing);


                worKsheeT = (Microsoft.Office.Interop.Excel.Worksheet)worKbooK.ActiveSheet;
                worKsheeT.Name = strNameSheet;

                for (int i = 1; i <= NameColumnKeyboard.Length; i++)
                {
                    worKsheeT.Cells[1, i] = NameColumnKeyboard[i - 1];
                }


                if (bdFull != null && bdFull.Rows.Count > 0)
                {
                    for (int i = 1; i <= bdFull.Rows.Count; i++)
                    {
                        string name = bdFull.Rows[i - 1][34].ToString();
                        string terminalid = bdFull.Rows[i - 1][20].ToString();

                        if (!string.IsNullOrEmpty(terminalid) && !string.IsNullOrEmpty(name))
                        {
                            #region tao sheet Terminal
                            // Class
                            worKsheeT.Cells[i + 1, 1] = "Keyboard";
                            // name                    
                            worKsheeT.Cells[i + 1, 2] = "Keyboard-" + name;
                            //system_name
                           // worKsheeT.Cells[i + 1, 3] = "Keyboard-" + name;
                            //serial_number
                            worKsheeT.Cells[i + 1, 4] = name;
                            //,"zlocation_area_ref"
                            string zlocation_area_ref = bdFull.Rows[i - 1][10].ToString();
                            worKsheeT.Cells[i + 1, 5] = zlocation_area_ref;
                            //active
                            worKsheeT.Cells[i + 1, 6] = "0";
                            #endregion
                        }
                    }
                }
                worKbooK.SaveAs(strNameFileNew); ;
                worKbooK.Close();
                excel.Quit();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
                throw;
            }
        }

        public void CreateTouchscreen(string strNameFileNew, string strNameSheet, DataTable bdFull)
        {
            Microsoft.Office.Interop.Excel.Application excel;
            Microsoft.Office.Interop.Excel.Workbook worKbooK;
            Microsoft.Office.Interop.Excel.Worksheet worKsheeT;
            Microsoft.Office.Interop.Excel.Range celLrangE;

            try
            {
                excel = new Microsoft.Office.Interop.Excel.Application();
                excel.Visible = false;
                excel.DisplayAlerts = false;
                worKbooK = excel.Workbooks.Add(Type.Missing);


                worKsheeT = (Microsoft.Office.Interop.Excel.Worksheet)worKbooK.ActiveSheet;
                worKsheeT.Name = strNameSheet;

                for (int i = 1; i <= NameColumnTouchscreen.Length; i++)
                {
                    worKsheeT.Cells[1, i] = NameColumnTouchscreen[i - 1];
                }


                if (bdFull != null && bdFull.Rows.Count > 0)
                {
                    for (int i = 1; i <= bdFull.Rows.Count; i++)
                    {
                        string name = bdFull.Rows[i - 1][35].ToString();
                        string terminalid = bdFull.Rows[i - 1][20].ToString();

                        if (!string.IsNullOrEmpty(terminalid) && !string.IsNullOrEmpty(name))
                        {
                            #region tao sheet Terminal
                            // Class
                            worKsheeT.Cells[i + 1, 1] = "Touchscreen";
                            // name                    
                            worKsheeT.Cells[i + 1, 2] = "Touchscreen-" + name;
                            //system_name
                           // worKsheeT.Cells[i + 1, 3] = "Touchscreen-" + name;
                            //serial_number
                            worKsheeT.Cells[i + 1, 4] = name;
                            //,"zlocation_area_ref"
                            string zlocation_area_ref = bdFull.Rows[i - 1][10].ToString();
                            worKsheeT.Cells[i + 1, 5] = zlocation_area_ref;
                            //active
                            worKsheeT.Cells[i + 1, 6] = "0";
                            #endregion
                        }
                    }
                }
                worKbooK.SaveAs(strNameFileNew); ;
                worKbooK.Close();
                excel.Quit();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
                throw;
            }
        }

        public void CreateCheckwin(string strNameFileNew, string strNameSheet, DataTable bdFull)
        {
            Microsoft.Office.Interop.Excel.Application excel;
            Microsoft.Office.Interop.Excel.Workbook worKbooK;
            Microsoft.Office.Interop.Excel.Worksheet worKsheeT;
            Microsoft.Office.Interop.Excel.Range celLrangE;

            try
            {
                excel = new Microsoft.Office.Interop.Excel.Application();
                excel.Visible = false;
                excel.DisplayAlerts = false;
                worKbooK = excel.Workbooks.Add(Type.Missing);


                worKsheeT = (Microsoft.Office.Interop.Excel.Worksheet)worKbooK.ActiveSheet;
                worKsheeT.Name = strNameSheet;

                for (int i = 1; i <= NameColumnCheckwin.Length; i++)
                {
                    worKsheeT.Cells[1, i] = NameColumnCheckwin[i - 1];
                }

                if (bdFull != null && bdFull.Rows.Count > 0)
                {
                    for (int i = 1; i <= bdFull.Rows.Count; i++)
                    {
                        string name = bdFull.Rows[i - 1][36].ToString();
                        string terminalid = bdFull.Rows[i - 1][20].ToString();

                        if (!string.IsNullOrEmpty(terminalid) && !string.IsNullOrEmpty(name))
                        {
                            #region tao sheet Terminal
                            // Class
                            worKsheeT.Cells[i + 1, 1] = "Checkwin";
                            // name 
                            worKsheeT.Cells[i + 1, 2] = "Checkwin-" + name;
                            //system_name
                           // worKsheeT.Cells[i + 1, 3] = "Checkwin-" + name;
                            //serial_number
                            worKsheeT.Cells[i + 1, 4] = name;
                            // ztype_ref (kieu lap dat)
                            string typeDeBan = bdFull.Rows[i - 1][26].ToString();
                            string typeGanTuong = bdFull.Rows[i - 1][27].ToString();
                            bool flag = false;
                            if (typeDeBan == "1")
                            {
                                worKsheeT.Cells[i + 1, 5] = "Để bàn";
                                flag = true;
                            }                            
                            if (typeGanTuong == "1")
                            {
                                worKsheeT.Cells[i + 1, 5] = "Gắn tường";
                                flag = true;
                            }
                            if(flag == false)
                                worKsheeT.Cells[i + 1, 5] = "EMPTY";
                            //,"zlocation_area_ref"
                            string zlocation_area_ref = bdFull.Rows[i - 1][10].ToString();
                            worKsheeT.Cells[i + 1, 6] = zlocation_area_ref;
                            //active
                            worKsheeT.Cells[i + 1, 7] = "0";
                            #endregion
                        }
                    }
                }
                worKbooK.SaveAs(strNameFileNew); ;
                worKbooK.Close();
                excel.Quit();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
                throw;
            }

        }

        public void CreateKho(string strNameFileNew, string strNameSheet, DataTable bdFull)
        {
            Microsoft.Office.Interop.Excel.Application excel;
            Microsoft.Office.Interop.Excel.Workbook worKbooK;
            Microsoft.Office.Interop.Excel.Worksheet worKsheeT;
            Microsoft.Office.Interop.Excel.Range celLrangE;

            try
            {
                excel = new Microsoft.Office.Interop.Excel.Application();
                excel.Visible = false;
                excel.DisplayAlerts = false;
                worKbooK = excel.Workbooks.Add(Type.Missing);


                worKsheeT = (Microsoft.Office.Interop.Excel.Worksheet)worKbooK.ActiveSheet;
                worKsheeT.Name = strNameSheet;

                for (int i = 1; i <= NameColumnKho.Length; i++)
                {
                    worKsheeT.Cells[1, i] = NameColumnKho[i - 1];
                }

                if (bdFull != null && bdFull.Rows.Count > 0)
                {
                    for (int i = 1; i <= bdFull.Rows.Count; i++)
                    {
                        string Tinh = bdFull.Rows[i - 1][10].ToString();
                        string tenkho = bdFull.Rows[i - 1][43].ToString();
                        string terminalid = bdFull.Rows[i - 1][20].ToString();

                        if (!string.IsNullOrEmpty(terminalid) && !string.IsNullOrEmpty(tenkho))
                        {
                            #region tao sheet Terminal
                            // Class
                            worKsheeT.Cells[i + 1, 1] = "Kho";
                            // name 
                            worKsheeT.Cells[i + 1, 2] = Tinh + "-" + tenkho;
                            //system_name
                            worKsheeT.Cells[i + 1, 3] = Tinh + "-" + tenkho;
                            //,"zlocation_area_ref"                       
                            worKsheeT.Cells[i + 1, 4] = Tinh;
                            //active
                            worKsheeT.Cells[i + 1, 5] = "0";
                            #endregion
                        }
                    }
                }
                worKbooK.SaveAs(strNameFileNew); ;
                worKbooK.Close();
                excel.Quit();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
                throw;
            }
        }

        public void CreateRealaPos_Bgt(string strNameFileNew, string strNameSheet, DataTable bdFull)
        {


            Microsoft.Office.Interop.Excel.Application excel;
            Microsoft.Office.Interop.Excel.Workbook worKbooK;
            Microsoft.Office.Interop.Excel.Worksheet worKsheeT;
            Microsoft.Office.Interop.Excel.Range celLrangE;

            try
            {
                excel = new Microsoft.Office.Interop.Excel.Application();
                excel.Visible = false;
                excel.DisplayAlerts = false;
                worKbooK = excel.Workbooks.Add(Type.Missing);


                worKsheeT = (Microsoft.Office.Interop.Excel.Worksheet)worKbooK.ActiveSheet;
                worKsheeT.Name = strNameSheet;

                for (int i = 1; i <= NameColumnRela_POS.Length; i++)
                {
                    worKsheeT.Cells[1, i] = NameColumnRela_POS[i - 1];
                }

                if (bdFull != null && bdFull.Rows.Count > 0)
                {
                    for (int i = 1; i <= bdFull.Rows.Count; i++)
                    {
                        string posid = bdFull.Rows[i - 1][0].ToString().Trim();
                        string bgtid = bdFull.Rows[i - 1][1].ToString().Trim();

                        List<string> lstPosId = GetDataID(posid, "zcode", "zpos");
                        List<string> lstBGTId = GetDataID(bgtid, "zcode", "zbgt_code");

                        if (!string.IsNullOrEmpty(posid) && lstPosId!= null && lstPosId.Count > 0)
                        {
                            #region tao sheet POS
                            // pos id  
                            worKsheeT.Cells[i + 1, 1] = lstPosId[0];
                            // class
                            worKsheeT.Cells[i + 1, 2] = "POS";
                            // bgt code 
                            worKsheeT.Cells[i + 1, 3] = (lstBGTId != null && lstBGTId.Count > 0) ? lstBGTId[0] : "";

                            #endregion
                        }

                    }
                }
                worKbooK.SaveAs(strNameFileNew); ;
                worKbooK.Close();
                excel.Quit();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
                throw;
            }
        }

        public void CreateRealaPos_Add(string strNameFileNew, string strNameSheet, DataTable bdFull)
        {
            Microsoft.Office.Interop.Excel.Application excel;
            Microsoft.Office.Interop.Excel.Workbook worKbooK;
            Microsoft.Office.Interop.Excel.Worksheet worKsheeT;
            Microsoft.Office.Interop.Excel.Range celLrangE;

            try
            {
                excel = new Microsoft.Office.Interop.Excel.Application();
                excel.Visible = false;
                excel.DisplayAlerts = false;
                worKbooK = excel.Workbooks.Add(Type.Missing);


                worKsheeT = (Microsoft.Office.Interop.Excel.Worksheet)worKbooK.ActiveSheet;
                worKsheeT.Name = strNameSheet;

                for (int i = 1; i <= NameColumnRela_POS_Add.Length; i++)
                {
                    worKsheeT.Cells[1, i] = NameColumnRela_POS_Add[i - 1];
                }


                if (bdFull != null && bdFull.Rows.Count > 0)
                {
                    for (int i = 1; i <= bdFull.Rows.Count; i++)
                    {
                        string posid = bdFull.Rows[i - 1][0].ToString().Trim();
                        string add_name = bdFull.Rows[i - 1][5].ToString().Trim();

                        List<string> lstPosId = GetDataID(posid, "zcode", "zpos");
                        List<string> lstAddId = GetDataID(add_name, "name", "nr");

                        if (!string.IsNullOrEmpty(posid) && lstPosId != null && lstPosId.Count > 0)
                        {
                            #region tao sheet POS
                            // add id  
                            worKsheeT.Cells[i + 1, 1] = (lstAddId != null && lstAddId.Count > 0) ? lstAddId[0] : "";
                            // class
                            worKsheeT.Cells[i + 1, 2] = "Address";
                            // pos
                            worKsheeT.Cells[i + 1, 3] = (lstPosId != null && lstPosId.Count > 0) ? lstPosId[0] : "";

                            #endregion
                        }

                    }
                }
                worKbooK.SaveAs(strNameFileNew); ;
                worKbooK.Close();
                excel.Quit();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
                throw;
            }
        }

        public void CreateRealaPos_Termi(string strNameFileNew, string strNameSheet, DataTable bdFull)
        {
            Microsoft.Office.Interop.Excel.Application excel;
            Microsoft.Office.Interop.Excel.Workbook worKbooK;
            Microsoft.Office.Interop.Excel.Worksheet worKsheeT;
            Microsoft.Office.Interop.Excel.Range celLrangE;

            try
            {
                excel = new Microsoft.Office.Interop.Excel.Application();
                excel.Visible = false;
                excel.DisplayAlerts = false;
                worKbooK = excel.Workbooks.Add(Type.Missing);


                worKsheeT = (Microsoft.Office.Interop.Excel.Worksheet)worKbooK.ActiveSheet;
                worKsheeT.Name = strNameSheet;

                for (int i = 1; i <= NameColumnRela_POS_Termi.Length; i++)
                {
                    worKsheeT.Cells[1, i] = NameColumnRela_POS_Termi[i - 1];
                }

                if (bdFull != null && bdFull.Rows.Count > 0)
                {
                    for (int i = 1; i <= bdFull.Rows.Count; i++)
                    {
                        string posid = bdFull.Rows[i - 1][0].ToString().Trim();
                        string terminalid = bdFull.Rows[i - 1][20].ToString();

                        List<string> lstPosId = GetDataID(posid, "zcode", "zpos");
                        List<string> lstTerminalId = GetDataID(terminalid, "zcode", "zterminal");

                        if (!string.IsNullOrEmpty(posid) && lstPosId!= null && lstPosId.Count> 0)
                        {
                            #region tao sheet POS
                            // add id  
                            worKsheeT.Cells[i + 1, 1] = (lstPosId != null && lstPosId.Count > 0) ? lstPosId[0] : "";
                            // class
                            worKsheeT.Cells[i + 1, 2] = "POS";
                            // pos
                            worKsheeT.Cells[i + 1, 3] = (lstTerminalId != null && lstTerminalId.Count > 0) ? lstTerminalId[0] : "";

                            #endregion
                        }

                    }
                }
                worKbooK.SaveAs(strNameFileNew); ;
                worKbooK.Close();
                excel.Quit();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
                throw;
            }

        }

        public void CreateRealaTermi_Other(string strNameFileNew, string strNameSheet, DataTable bdFull)
        {

            Microsoft.Office.Interop.Excel.Application excel;
            Microsoft.Office.Interop.Excel.Workbook worKbooK;
            Microsoft.Office.Interop.Excel.Worksheet worKsheeT;
            Microsoft.Office.Interop.Excel.Range celLrangE;

            try
            {
                excel = new Microsoft.Office.Interop.Excel.Application();
                excel.Visible = false;
                excel.DisplayAlerts = false;
                worKbooK = excel.Workbooks.Add(Type.Missing);


                worKsheeT = (Microsoft.Office.Interop.Excel.Worksheet)worKbooK.ActiveSheet;
                worKsheeT.Name = strNameSheet;

                for (int i = 1; i <= NameColumnRela_Terminal_Other.Length; i++)
                {
                    worKsheeT.Cells[1, i] = NameColumnRela_Terminal_Other[i - 1];
                }


                if (bdFull != null && bdFull.Rows.Count > 0)
                {
                    for (int i = 1; i <= bdFull.Rows.Count; i++)
                    {
                        string terminalid = bdFull.Rows[i - 1][20].ToString();
                        string sim3g = bdFull.Rows[i - 1][16].ToString();
                        string checkwin = bdFull.Rows[i - 1][36].ToString();
                        string router = bdFull.Rows[i - 1][28].ToString();
                        string ups = bdFull.Rows[i - 1][29].ToString();
                        string cpu = bdFull.Rows[i - 1][30].ToString();
                        string cdu = bdFull.Rows[i - 1][31].ToString();
                        string printer = bdFull.Rows[i - 1][32].ToString();
                        string scanner = bdFull.Rows[i - 1][33].ToString();
                        string keyboard = bdFull.Rows[i - 1][34].ToString();
                        string touchscreen = bdFull.Rows[i - 1][35].ToString();

                        List<string> lstTerminalId = GetDataID(terminalid, "zcode", "zterminal");

                        if (lstTerminalId != null && lstTerminalId.Count > 0)
                        {
                            List<string> lstSim3G = GetDataID(sim3g, "serial_number", "nr");
                            List<string> lstcheckwin = GetDataID(checkwin, "serial_number", "nr");
                            List<string> lstrouter = GetDataID(router, "serial_number", "nr");
                            List<string> lstups = GetDataID(ups, "serial_number", "nr");
                            List<string> lstcpu = GetDataID(cpu, "serial_number", "nr");
                            List<string> lstcdu = GetDataID(cdu, "serial_number", "nr");
                            List<string> lstprinter = GetDataID(printer, "serial_number", "nr");
                            List<string> lstscanner = GetDataID(scanner, "serial_number", "nr");
                            List<string> lstkeyboard = GetDataID(keyboard, "serial_number", "nr");
                            List<string> lsttouchscreen = GetDataID(touchscreen, "serial_number", "nr");

                            #region tao sheet POS
                            // terminal  
                            worKsheeT.Cells[i + 1, 1] = (lstTerminalId != null && lstTerminalId.Count > 0) ? lstTerminalId[0] : "";
                            // class
                            worKsheeT.Cells[i + 1, 2] = "Terminal";
                            // psim3g
                            worKsheeT.Cells[i + 1, 3] = (lstSim3G != null && lstSim3G.Count > 0) ? lstSim3G[0] : "";
                            // checkwin
                            worKsheeT.Cells[i + 1, 4] = (lstcheckwin != null && lstcheckwin.Count > 0) ? lstcheckwin[0] : "";
                            //router
                            worKsheeT.Cells[i + 1, 5] = (lstrouter != null && lstrouter.Count > 0) ? lstrouter[0] : "";
                            //ups
                            worKsheeT.Cells[i + 1, 6] = (lstups != null && lstups.Count > 0) ? lstups[0] : "";
                            //cpu
                            worKsheeT.Cells[i + 1, 7] = (lstcpu != null && lstcpu.Count > 0) ? lstcpu[0] : "";
                            //cdu
                            worKsheeT.Cells[i + 1, 8] = (lstcdu != null && lstcdu.Count > 0) ? lstcdu[0] : "";
                            //printer
                            worKsheeT.Cells[i + 1, 9] = (lstprinter != null && lstprinter.Count > 0) ? lstprinter[0] : "";
                            //scanner
                            worKsheeT.Cells[i + 1, 10] = (lstscanner != null && lstscanner.Count > 0) ? lstscanner[0] : "";
                            //keyboard
                            worKsheeT.Cells[i + 1, 11] = (lstkeyboard != null && lstkeyboard.Count > 0) ? lstkeyboard[0] : "";
                            //touch
                            worKsheeT.Cells[i + 1, 12] = (lsttouchscreen != null && lsttouchscreen.Count > 0) ? lsttouchscreen[0] : "";

                            #endregion
                        }

                    }
                }
                worKbooK.SaveAs(strNameFileNew); ;
                worKbooK.Close();
                excel.Quit();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
                throw;
            }

        }

        public void CreateRealaKho_Other(string strNameFileNew, string strNameSheet, DataTable bdFull)
        {
            Microsoft.Office.Interop.Excel.Application excel;
            Microsoft.Office.Interop.Excel.Workbook worKbooK;
            Microsoft.Office.Interop.Excel.Worksheet worKsheeT;
            Microsoft.Office.Interop.Excel.Range celLrangE;

            try
            {
                excel = new Microsoft.Office.Interop.Excel.Application();
                excel.Visible = false;
                excel.DisplayAlerts = false;
                worKbooK = excel.Workbooks.Add(Type.Missing);


                worKsheeT = (Microsoft.Office.Interop.Excel.Worksheet)worKbooK.ActiveSheet;
                worKsheeT.Name = strNameSheet;

                for (int i = 1; i <= NameColumnRela_kho_other.Length; i++)
                {
                    worKsheeT.Cells[1, i] = NameColumnRela_kho_other[i - 1];
                }

                if (bdFull != null && bdFull.Rows.Count > 0)
                {
                    for (int i = 1; i <= bdFull.Rows.Count; i++)
                    {
                        string tinh = bdFull.Rows[i - 1][10].ToString();
                        string khoname = bdFull.Rows[i - 1][43].ToString();
                        string terminalid = bdFull.Rows[i - 1][20].ToString();
                        string sim3g = bdFull.Rows[i - 1][16].ToString();
                        string checkwin = bdFull.Rows[i - 1][36].ToString();
                        string router = bdFull.Rows[i - 1][28].ToString();
                        string ups = bdFull.Rows[i - 1][29].ToString();
                        string cpu = bdFull.Rows[i - 1][30].ToString();
                        string cdu = bdFull.Rows[i - 1][31].ToString();
                        string printer = bdFull.Rows[i - 1][32].ToString();
                        string scanner = bdFull.Rows[i - 1][33].ToString();
                        string keyboard = bdFull.Rows[i - 1][34].ToString();
                        string touchscreen = bdFull.Rows[i - 1][35].ToString();

                        List<string> lstKho = GetDataID(tinh + "-" + khoname, "name", "nr");

                        if (lstKho != null && lstKho.Count > 0)
                        {
                            List<string> lstTerminalId = GetDataID(terminalid, "zcode", "zterminal");
                            List<string> lstSim3G = GetDataID(sim3g, "serial_number", "nr");
                            List<string> lstcheckwin = GetDataID(checkwin, "serial_number", "nr");
                            List<string> lstrouter = GetDataID(router, "serial_number", "nr");
                            List<string> lstups = GetDataID(ups, "serial_number", "nr");
                            List<string> lstcpu = GetDataID(cpu, "serial_number", "nr");
                            List<string> lstcdu = GetDataID(cdu, "serial_number", "nr");
                            List<string> lstprinter = GetDataID(printer, "serial_number", "nr");
                            List<string> lstscanner = GetDataID(scanner, "serial_number", "nr");
                            List<string> lstkeyboard = GetDataID(keyboard, "serial_number", "nr");
                            List<string> lsttouchscreen = GetDataID(touchscreen, "serial_number", "nr");

                            #region tao sheet POS
                            //kho
                            worKsheeT.Cells[i + 1, 1] = (lstKho != null && lstKho.Count > 0) ? lstKho[0] : "";
                            // class
                            worKsheeT.Cells[i + 1, 2] = "Kho";
                            // terminal  
                            worKsheeT.Cells[i + 1, 3] = (lstTerminalId != null && lstTerminalId.Count > 0) ? lstTerminalId[0] : "";
                            // psim3g
                            worKsheeT.Cells[i + 1, 4] = (lstSim3G != null && lstSim3G.Count > 0) ? lstSim3G[0] : "";
                            // checkwin
                            worKsheeT.Cells[i + 1, 5] = (lstcheckwin != null && lstcheckwin.Count > 0) ? lstcheckwin[0] : "";
                            //router
                            worKsheeT.Cells[i + 1, 6] = (lstrouter != null && lstrouter.Count > 0) ? lstrouter[0] : "";
                            //ups
                            worKsheeT.Cells[i + 1, 7] = (lstups != null && lstups.Count > 0) ? lstups[0] : "";
                            //cpu
                            worKsheeT.Cells[i + 1, 8] = (lstcpu != null && lstcpu.Count > 0) ? lstcpu[0] : "";
                            //cdu
                            worKsheeT.Cells[i + 1, 9] = (lstcdu != null && lstcdu.Count > 0) ? lstcdu[0] : "";
                            //printer
                            worKsheeT.Cells[i + 1, 10] = (lstprinter != null && lstprinter.Count > 0) ? lstprinter[0] : "";
                            //scanner
                            worKsheeT.Cells[i + 1, 11] = (lstscanner != null && lstscanner.Count > 0) ? lstscanner[0] : "";
                            //keyboard
                            worKsheeT.Cells[i + 1, 12] = (lstkeyboard != null && lstkeyboard.Count > 0) ? lstkeyboard[0] : "";
                            //touch
                            worKsheeT.Cells[i + 1, 13] = (lsttouchscreen != null && lsttouchscreen.Count > 0) ? lsttouchscreen[0] : "";

                            #endregion
                        }

                    }
                }

                worKbooK.SaveAs(strNameFileNew); ;
                worKbooK.Close();
                excel.Quit();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
                throw;
            }

        }



        public static List<string> GetDataID(string namevalue, string name_attribute, string nameobject)
        {
            try
            {
                List<string> lst = new List<string>();
                string[] attr = { "id", name_attribute };
                XDocument xml = new XDocument();
                string UDSObj = ca.doSelect(SID, nameobject, name_attribute + " = '" + namevalue + "'", -1, attr);
                xml = XDocument.Parse(UDSObj);

                foreach (XElement element in xml.Descendants("UDSObject"))
                {
                    foreach (XElement EAttr in element.Descendants("Attribute"))
                    {
                        lst.Add(EAttr.Element("AttrValue").Value);
                    }
                }
                return lst;
            }
            catch (Exception ex)
            {
                return null;
            }
        }

        private DataTable tbReadEx(string strFilePath, string strSheet)
        {
            DataTable tblDataExcel = new DataTable();
            //string strConn = string.Format("Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties=\"Excel 12.0 Xml;HDR=YES;IMEX=1;TypeGuessRows=0;ImportMixedTypes=Text\"", strFilePath);
            string strConn = string.Format("Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties=\"Excel 12.0 Xml;HDR=YES;IMEX=1;TypeGuessRows=0\"", strFilePath);
            using (OleDbConnection dbConnection = new OleDbConnection(strConn))
            {

                try
                {
                    using (OleDbDataAdapter dbAdapter = new OleDbDataAdapter("SELECT * FROM [" + strSheet + "$]", dbConnection)) //rename sheet if required!
                        dbAdapter.Fill(tblDataExcel);
                }
                catch (Exception)
                {
                    //throw;
                    flag = false;
                    return null;
                }

            }
            return tblDataExcel;


        }

        private void btnTachFile_Click(object sender, EventArgs e)
        {
            string Folder_New = string.Empty;
            string nameFileNew = string.Empty;
            string strSheetExport = string.Empty;
            string sheetImport = string.Empty;
            string Time = DateTime.Now.ToLongTimeString().Replace(":", "_").Replace(" ", "_");
            DataTable dbFull, dbBGT, dbSim3G;

            Folder_New = txtFolderLog.Text + "\\" + Time;
            if (!System.IO.Directory.Exists(Folder_New))
            {
                System.IO.Directory.CreateDirectory(Folder_New);
            }

            if (!CheckFileTemplate(txtFile.Text))
            {
                MessageBox.Show("Kiểm tra lại quy cách đặt tên sheet theo đúng mẫu quy định.");
                WriteOLog("Kiểm tra lại quy cách đặt tên sheet theo đúng mẫu quy định.");
                return;
            }

            try
            {
                // Đọc dữ liệu từ sheet FullCI
                sheetImport = "FullCI";
                dbFull = tbReadEx(txtFile.Text, sheetImport);

                // Đọc dữ liệu từ sheet BGT Code
                sheetImport = "BGT";
                dbBGT = tbReadEx(txtFile.Text, sheetImport);

                //Đọc dữ liệu từ sheet Sim3G
                sheetImport = "Sim3G";
                dbSim3G = tbReadEx(txtFile.Text, sheetImport);

                if (!CheckCountColumn(dbFull, dbBGT, dbSim3G))
                {
                    MessageBox.Show("Kiểm tra lại file tách dữ liệu theo đúng mẫu quy định.");
                    WriteOLog("Kiểm tra lại file tách dữ liệu theo đúng mẫu quy định.");
                    return;
                }               
            }
            catch (Exception)
            {
                MessageBox.Show("Kiểm tra lại quy cách đặt tên của file và sheet");
                WriteOLog("Kiểm tra lại quy cách đặt tên của file và sheet");
                return;
                throw;
            }

            try
            {
               // MessageBox.Show("bat dau tach POS");
                #region tách file POS
                nameFileNew = Folder_New + "\\POS_" + Time + ".xlsx";
                strSheetExport = "POS";
                CreatePOS(nameFileNew, strSheetExport, dbFull);
                #endregion
                //MessageBox.Show("bat dau tach POS");
                #region Tách file Terminal
                nameFileNew = Folder_New + "\\Terminal_" + Time + ".xlsx";
                strSheetExport = "Terminal";
                CreateTerminal(nameFileNew, strSheetExport, dbFull);
                #endregion

                #region Tách file BGT
                nameFileNew = Folder_New + "\\BGT_" + Time + ".xlsx";
                strSheetExport = "BGT";
                CreateBGTCode(nameFileNew, strSheetExport, dbBGT);
                #endregion

                #region Tách file Sim3G
                nameFileNew = Folder_New + "\\Sim3G_" + Time + ".xlsx";
                strSheetExport = "Sim3G";
                CreateSim3G(nameFileNew, strSheetExport, dbSim3G);
                #endregion

                #region tách file Location
                nameFileNew = Folder_New + "\\Loc_" + Time + ".xlsx";
                strSheetExport = "Location";
                CreateLocation(nameFileNew, strSheetExport, dbFull);
                #endregion

                #region tách file router
                nameFileNew = Folder_New + "\\Router_" + Time + ".xlsx";
                strSheetExport = "Router";
                CreateRouter(nameFileNew, strSheetExport, dbFull);
                #endregion

                #region Tách file UPS
                nameFileNew = Folder_New + "\\UPS_" + Time + ".xlsx";
                strSheetExport = "UPS";
                CreateUPS(nameFileNew, strSheetExport, dbFull);
                #endregion

                #region Tách file CPU
                nameFileNew = Folder_New + "\\CPU_" + Time + ".xlsx";
                strSheetExport = "CPU";
                CreateCPU(nameFileNew, strSheetExport, dbFull);
                #endregion

                #region Tách file CDU
                nameFileNew = Folder_New + "\\CDU_" + Time + ".xlsx";
                strSheetExport = "CDU";
                CreateCDU(nameFileNew, strSheetExport, dbFull);
                #endregion

                #region Tách file Printer
                nameFileNew = Folder_New + "\\Printer_" + Time + ".xlsx";
                strSheetExport = "Printer";
                CreatePrinter(nameFileNew, strSheetExport, dbFull);
                #endregion

                #region Tách file Scanner
                nameFileNew = Folder_New + "\\Scanner" + Time + ".xlsx";
                strSheetExport = "Scanner";
                CreateScanner(nameFileNew, strSheetExport, dbFull);
                #endregion

                #region Tách file Keyboard
                nameFileNew = Folder_New + "\\Keyboard" + Time + ".xlsx";
                strSheetExport = "Keyboard";
                CreateKeyboard(nameFileNew, strSheetExport, dbFull);
                #endregion

                #region Tách file Touchscreen
                nameFileNew = Folder_New + "\\Touchscreen" + Time + ".xlsx";
                strSheetExport = "Touchscreen";
                CreateTouchscreen(nameFileNew, strSheetExport, dbFull);
                #endregion

                #region Tách file Checkwin
                nameFileNew = Folder_New + "\\Checkwin" + Time + ".xlsx";
                strSheetExport = "Checkwin";
                CreateCheckwin(nameFileNew, strSheetExport, dbFull);
                #endregion

                #region Tách file Kho
                nameFileNew = Folder_New + "\\Kho_" + Time + ".xlsx";
                strSheetExport = "Kho";
                CreateKho(nameFileNew, strSheetExport, dbFull);//(nameFileNew, strSheetExport, dbFull);
                #endregion

               
                MessageBox.Show("Tách file dữ liệu thành công. Kiểm tra dữ liệu đã tách tại thư mục: \n" + Folder_New);
                WriteOLog("Tách file dữ liệu thành công. Kiểm tra dữ liệu đã tách tại thư mục: \n" + Folder_New);
               
            }
            catch (Exception)
            {
                MessageBox.Show("Có lỗi trong quá trình tách dữ liệu. Kiểm tra lại log hệ thống");
                WriteOLog("Có lỗi trong quá trình tách dữ liệu. Kiểm tra lại log hệ thống" + Folder_New);
                return;
                throw;
            }

        }

        /// <summary>
        /// Chọn file 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void BrowseFile_Click(object sender, EventArgs e)
        {
            txtThongBao.Text = string.Empty;
            Stream myStream = null;
            OpenFileDialog openFileDialog1 = new OpenFileDialog();

            openFileDialog1.InitialDirectory = "c:\\";
            openFileDialog1.Filter = "Excel files (*.xls)|*.xls|All files (*.*)|*.*";
            openFileDialog1.FilterIndex = 2;
            openFileDialog1.RestoreDirectory = true;
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                txtFile.Text = openFileDialog1.FileName;
                bindDataSheetExcel(txtFile.Text);
            }
        }

        /// <summary>
        /// Load sheet into comboname
        /// </summary>
        /// <param name="strPathFile"></param>
        private void bindDataSheetExcel(string strPathFile)
        {
            DataTable tblDataExcel = new DataTable();
            DataTable tblSheet = new DataTable("TABLE_NAME");
            DataRow dRowSheets;
            tblSheet.Columns.Add("TABLE_NAME");

            try
            {
                string strConn = string.Format("Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties=\"Excel 12.0 Xml;HDR=YES;IMEX=1;TypeGuessRows=0;ImportMixedTypes=Text\"", strPathFile);
                // string strConn = string.Format("Provider=Microsoft.Jet.OLEDB.4.0;Data Source={0}; Extended Properties=\"Excel 8.0; HDR=Yes; IMEX=1\"", strPathFile);
                using (OleDbConnection dbConnection = new OleDbConnection(strConn))
                {
                    dbConnection.Open();
                    DataTable tblResult = dbConnection.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                    foreach (DataRow dRowSheet in tblResult.Rows)
                    {
                        dRowSheets = tblSheet.NewRow();
                        dRowSheets[0] = dRowSheet["TABLE_NAME"].ToString().Replace("$", "");
                        tblSheet.Rows.Add(dRowSheets);
                    }

                }
                cmbSheet.DataSource = tblSheet;
                cmbSheet.DisplayMember = "TABLE_NAME";
                cmbSheet.ValueMember = "TABLE_NAME";
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString().Substring(0, 200) + "...", " Message Error");
                WriteOLog(ex.ToString().Substring(0, 200) + "...");
                return;
            }
        }

        /// <summary>
        /// Kiểm tra, sau khi chạy GRloader, hệ thống có gen ra file lỗi hay không?
        /// </summary>
        /// <param name="pathfile"></param>
        /// <returns></returns>
        public bool IsErr(string pathfile_err)
        {
            if (File.Exists(pathfile_err))
                return true;
            return false;

        }

        /// <summary>
        /// Import CI
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void tbnImportCI_Click(object sender, EventArgs e)
        {
            string pathFile = txtFile.Text;
            string namsheet = cmbSheet.Text;

            string FolderName = Path.GetDirectoryName(pathFile);

            string fileName = Path.GetFileNameWithoutExtension(pathFile);

            string filenameerr = fileName + "_err.xml";

            if (IsErr(FolderName + "\\" + filenameerr))// neu da co file err
            {
                // MessageBox.Show("Đã có file err.Hay xoa truoc khi thuc hien");
                string fileNamecopyed = System.IO.File.GetLastWriteTime(FolderName + "\\" + filenameerr).ToString();
                fileNamecopyed = fileNamecopyed.Replace(@"\", "_").Replace(@"/", "_").Replace(":", "_");
                //cut file sang folder History
                string FoldeHistory = FolderName + "\\History";
                if (!System.IO.Directory.Exists(FoldeHistory))
                {
                    System.IO.Directory.CreateDirectory(FoldeHistory);
                }
                try
                {
                    System.IO.File.Copy(FolderName + "\\" + filenameerr, FolderName + "\\History\\" + fileName + "_err_" + fileNamecopyed + ".xml", true);
                }
                catch (Exception ex)
                {
                    // ghi log 
                    MessageBox.Show("Copy file history không thành công" + ex);
                    WriteOLog("Copy file history không thành công");
                }
                try
                {
                    File.Delete(FolderName + "\\" + filenameerr);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Xóa file err của lần chạy trước không thành công");
                    WriteOLog("Xóa file err của lần chạy trước không thành công");
                    throw;
                }

            }
            string value = commandline_importci(txtFolderGrloader.Text, "10.33.3.124", txtUser.Text, txtPass.Text, pathFile, namsheet);
            MessageBox.Show(value);

            // ExecuteCommandSync(@"java -cp D:\SetUp -jar D:\SetUp\java\lib\GRLoader.jar -N D:\SetUp -u adminsrv -s 10.33.3.124 -i " + pathFile + @" -a -n -E -u adminsrv -p fiss@123 -s http://10.33.3.124:80 -sss " + namsheet);
            // Kiem tra co file err xuat ra hay khong. Neu co: thong bao GRLoader chay loi. Neu khong: thong bao thành cong
            if (IsErr(FolderName + "\\" + filenameerr))
            {
                MessageBox.Show("GRloader lỗi. Kiểm tra log tại file " + FolderName + "\\" + filenameerr);
                WriteOLog("GRloader lỗi. Kiểm tra log tại file " + FolderName + "\\" + filenameerr);

            }
            else
            {
                MessageBox.Show("Import CI thành công");
                WriteOLog("Import CI thành công");
            }

        }

        /// <summary>
        /// Run Grloader - File Excel
        /// </summary>
        /// <param name="grloader_folder"></param>
        /// <param name="hostname"></param>
        /// <param name="user"></param>
        /// <param name="pass"></param>
        /// <param name="pathfile"></param>
        /// <param name="sheetname"></param>
        /// <returns></returns>
        public string commandline_importci(string grloader_folder, string hostname, string user, string pass, string pathfile, string sheetname)
        {
            string command = string.Empty;
            command = @"java -cp " + grloader_folder + " -jar " + grloader_folder + @"\java\lib\GRLoader.jar -N " + grloader_folder + " -u " + user + " -s " + hostname + " -i " + pathfile + @" -a -n -E -u " + user + " -p " + pass + " -s http://" + hostname + ":80 -sss " + sheetname;
            return ExecuteCommandSync(command);
        }

        /// <summary>
        /// Run Gloader - File xml
        /// </summary>
        /// <param name="grloader_folder"></param>
        /// <param name="hostname"></param>
        /// <param name="user"></param>
        /// <param name="pass"></param>
        /// <param name="pathfile"></param>
        /// <returns></returns>
        public string commandline_importRelation(string grloader_folder, string hostname, string user, string pass, string pathfile)
        {
            string command = string.Empty;
            command = @"java -cp " + grloader_folder + " -jar " + grloader_folder + @"\java\lib\GRLoader.jar -N " + grloader_folder + " -u " + user + " -s " + hostname + " -i " + pathfile + @" -a -n -E -u " + user + " -p " + pass + " -s http://" + hostname + ":80";
            return ExecuteCommandSync(command);
        }

        /// <summary>
        /// Import dữ liệu: Chạy GRloader
        /// </summary>
        /// <param name="command"></param>
        public string ExecuteCommandSync(object command)
        {
            string value = string.Empty;
            try
            {

                System.Diagnostics.ProcessStartInfo procStartInfo =
                    new System.Diagnostics.ProcessStartInfo("cmd", "/c " + command);

                procStartInfo.RedirectStandardOutput = true;
                procStartInfo.UseShellExecute = false;
                // Do not create the black window.
                procStartInfo.CreateNoWindow = true;
                // Now we create a process, assign its ProcessStartInfo and start it
                System.Diagnostics.Process proc = new System.Diagnostics.Process();
                proc.StartInfo = procStartInfo;
                proc.Start();
                // Get the output into a string
                value = proc.StandardOutput.ReadToEnd();
            }
            catch (Exception objException)
            {
                value = "error";
            }
            return value;
        }

        private void btnRelationship_Click(object sender, EventArgs e)
        {
            string pathFile = txtFile.Text;
            string namsheet = cmbSheet.Text;

            string FolderName = Path.GetDirectoryName(pathFile);

            string fileName = Path.GetFileNameWithoutExtension(pathFile);

            string filenameerr = FolderName + "\\" + fileName + "_err.xml";

            // chayj grloader import Relationship lan 1
            string value = commandline_importci(txtFolderGrloader.Text, "10.33.3.124", txtUser.Text, txtPass.Text, pathFile, namsheet);

            if (IsErr(filenameerr))// neu da co file err
            {
                // khi co file err. Doc file err.xml
                FileStream fs = new FileStream(filenameerr, FileMode.Open);
                StreamReader rd = new StreamReader(fs, Encoding.UTF8);
                String mess_xml = rd.ReadToEnd();// ReadLine() chỉ đọc 1 dòng đầu thoy, ReadToEnd là đọc hết

                mess_xml = mess_xml.Replace("<name>", "<id>").Replace("</name>", "</id>");
                rd.Close();
                // Xóa file cũ 
                File.Delete(filenameerr);
                // ghi noi dung vao file moi, giống hệ tên file cũ
                FileStream fs2 = new FileStream(filenameerr, FileMode.OpenOrCreate);//Tạo file mới tên là test.txt            
                StreamWriter sWriter = new StreamWriter(fs2, Encoding.UTF8);//fs là 1 FileStream 
                sWriter.Write(mess_xml);
                sWriter.Flush();
                fs2.Close();
                // import Grloader trên file xml này
                string value2 = commandline_importRelation(txtFolderGrloader.Text, "10.33.3.124", txtUser.Text, txtPass.Text, filenameerr);


                string fileName_err2 = Path.GetFileNameWithoutExtension(filenameerr);

                string fileName_new = FolderName + "\\" + fileName_err2 + "_err.xml";

                // Kiem tra co file err xuat ra hay khong. Neu co: thong bao GRLoader chay loi. Neu khong: thong bao thành cong
                if (IsErr(fileName_new))
                {
                    MessageBox.Show("Import Relationship không thành công. Kiểm tra log tại file " + fileName);
                    WriteOLog("Import Relationship không thành công. Kiểm tra log tại file " + fileName);
                    File.Delete(filenameerr);
                }
                else
                {
                    MessageBox.Show("Import Relationship thành công");
                    WriteOLog("Import Relationship thành công");
                    File.Delete(filenameerr);
                }

            }
            else
            {
                MessageBox.Show("Import Relationship thành công");
                WriteOLog("Import Relationship thành công");
            }
        }

        //public void inserRe(string fileerr)
        //{
        //    ExecuteCommandSync(@"java -cp D:\SetUp -jar D:\SetUp\java\lib\GRLoader.jar -N D:\SetUp -u adminsrv -s 10.33.3.124 -i " + fileerr + @" -a -n -E -u adminsrv -p fiss@123 -s http://10.33.3.124:80");

        //    string fileName = Path.GetFileNameWithoutExtension(fileerr);

        //    fileName = fileName + "_err.xml";

        //    // Kiem tra co file err xuat ra hay khong. Neu co: thong bao GRLoader chay loi. Neu khong: thong bao thành cong
        //    if (IsErr(fileName))
        //    {
        //        MessageBox.Show("GRloader lỗi. Kiểm tra log tại file");

        //    }
        //    else
        //        MessageBox.Show("GRloader thành công");
        //}


        private void button1_Click(object sender, EventArgs e)
        {
            string Folder_New = string.Empty;
            string nameFileNew = string.Empty;
            string strSheetExport = string.Empty;
            string sheetImport = string.Empty;
            string Time = DateTime.Now.ToLongTimeString().Replace(":", "_").Replace(" ", "_");
            DataTable dbFull;

            Folder_New = txtFolderLog.Text + "\\" + Time;
            if (!System.IO.Directory.Exists(Folder_New))
            {
                System.IO.Directory.CreateDirectory(Folder_New);
            }

            try
            {
                // Đọc dữ liệu từ sheet FullCI
                sheetImport = "FullCI";
                dbFull = tbReadEx(txtFile.Text, sheetImport);

            }
            catch (Exception)
            {
                MessageBox.Show("Kiểm tra lại quy cách đặt tên của file và sheet");
                WriteOLog("Kiểm tra lại quy cách đặt tên của file và sheet");
                return;
                throw;
            }

            try
            {


                #region Tách file Relationship Terminal - Other
                nameFileNew = Folder_New + "\\Kho_" + Time + ".xlsx";
                strSheetExport = "Kho";
                CreateRealaKho_Other(nameFileNew, strSheetExport, dbFull);//(nameFileNew, strSheetExport, dbFull);
                #endregion

                MessageBox.Show("Tách file dữ liệu thành công. Kiểm tra dữ liệu đã tách tại thư mục: \n" + Folder_New);
                WriteOLog("Tách file dữ liệu thành công. Kiểm tra dữ liệu đã tách tại thư mục: \n" + Folder_New);
                txtThongBao.Text = "Tách file dữ liệu thành công. Kiểm tra dữ liệu đã tách tại thư mục: \n" + Folder_New;
                txtThongBao.Visible = true;
            }
            catch (Exception)
            {
                MessageBox.Show("Có lỗi trong quá trình tách dữ liệu. Kiểm tra lại log hệ thống");
                WriteOLog("Tách file dữ liệu thành công. Kiểm tra dữ liệu đã tách tại thư mục: \n" + Folder_New);
                return;
                throw;
            }

        }


        /// <summary>
        /// Kieemr tra file dau vao trước khi tách dữ liệu đã đúng thông tin sheet hay chưa?
        /// </summary>
        /// <returns></returns>
        public bool CheckFileTemplate(string strPathFile)
        {
            DataTable tblDataExcel = new DataTable();
            DataTable tblSheet = new DataTable("TABLE_NAME");
            tblSheet.Columns.Add("TABLE_NAME");

            string[] listSheet = new string[100];

            try
            {
                string strConn = string.Format("Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties=\"Excel 12.0 Xml;HDR=YES;IMEX=1;TypeGuessRows=0;ImportMixedTypes=Text\"", strPathFile);
                // string strConn = string.Format("Provider=Microsoft.Jet.OLEDB.4.0;Data Source={0}; Extended Properties=\"Excel 8.0; HDR=Yes; IMEX=1\"", strPathFile);
                using (OleDbConnection dbConnection = new OleDbConnection(strConn))
                {
                    dbConnection.Open();
                    DataTable tblResult = dbConnection.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                    int i = 0;
                    foreach (DataRow dRowSheet in tblResult.Rows)
                    {
                        string namesheet = dRowSheet["TABLE_NAME"].ToString().Replace("$", "");
                        listSheet[i] = namesheet;
                        i++;
                    }

                    foreach (string check in NameSheetTemplate)
                    {
                        if (!CheckInList(listSheet, check))
                            return false;
                    }
                }

            }
            catch (Exception ex)
            {

                return false;
            }
            return true;
        }

        /// <summary>
        ///  Kiểm tra namecheck có tồn tại trong danh sách list hay không?
        /// </summary>
        /// <param name="list"></param>
        /// <param name="namecheck"></param>
        /// <returns></returns>
        public bool CheckInList(string[] list, string namecheck)
        {
            foreach (string str in list)
            {
                if (namecheck.Equals(str))
                    return true;
            }
            return false;
        }

        /// <summary>
        /// Kiểm tra file trước khi tách dữ liệu đã đủ số lượng các cột yêu cầu hay chưa?
        /// </summary>
        /// <param name="dbFull"></param>
        /// <param name="dbBgt"></param>
        /// <param name="dbSim3G"></param>
        /// <returns></returns>
        public bool CheckCountColumn(DataTable dbFull, DataTable dbBgt, DataTable dbSim3G)
        {
            if (dbFull == null || dbFull.Columns.Count < 46)
                return false;

            if (dbFull == null || dbBgt.Columns.Count < 7)
                return false;

            if (dbFull == null || dbSim3G.Columns.Count < 3)
                return false;
            return true;
        }

        private void btnTachRela_Click(object sender, EventArgs e)
        {
            string Folder_New = string.Empty;
            string nameFileNew = string.Empty;
            string strSheetExport = string.Empty;
            string sheetImport = string.Empty;
            string Time = DateTime.Now.ToLongTimeString().Replace(":", "_").Replace(" ", "_");
            DataTable dbFull, dbBGT, dbSim3G;

            Folder_New = txtFolderLog.Text + "\\" + Time;
            if (!System.IO.Directory.Exists(Folder_New))
            {
                System.IO.Directory.CreateDirectory(Folder_New);
            }

            if (!CheckFileTemplate(txtFile.Text))
            {
                MessageBox.Show("Kiểm tra lại quy cách đặt tên sheet theo đúng mẫu quy định.");
                WriteOLog("Kiểm tra lại quy cách đặt tên sheet theo đúng mẫu quy định.");
                return;
            }

            try
            {
                // Đọc dữ liệu từ sheet FullCI
                sheetImport = "FullCI";
                dbFull = tbReadEx(txtFile.Text, sheetImport);

                // Đọc dữ liệu từ sheet BGT Code
                sheetImport = "BGT";
                dbBGT = tbReadEx(txtFile.Text, sheetImport);

                //Đọc dữ liệu từ sheet Sim3G
                sheetImport = "Sim3G";
                dbSim3G = tbReadEx(txtFile.Text, sheetImport);

                if (!CheckCountColumn(dbFull, dbBGT, dbSim3G))
                {
                    MessageBox.Show("Kiểm tra lại file tách dữ liệu theo đúng mẫu quy định.");
                    WriteOLog("Kiểm tra lại file tách dữ liệu theo đúng mẫu quy định.");
                    return;
                }
            }
            catch (Exception)
            {
                MessageBox.Show("Kiểm tra lại quy cách đặt tên của file và sheet");
                WriteOLog("Kiểm tra lại quy cách đặt tên của file và sheet");
                return;
                throw;
            }

            try
            {
                #region Tách file Relationship Pos - bgt
                nameFileNew = Folder_New + "\\Relationship_pos_bgt" + Time + ".xlsx";
                strSheetExport = "pos_bgt";
                CreateRealaPos_Bgt(nameFileNew, strSheetExport, dbFull);
                #endregion

                #region Tách file Relationship Pos - bgt
                nameFileNew = Folder_New + "\\Relationship_pos_add" + Time + ".xlsx";
                strSheetExport = "pos_add";
                CreateRealaPos_Add(nameFileNew, strSheetExport, dbFull);
                #endregion

                #region Tách file Relationship Pos - terminal
                nameFileNew = Folder_New + "\\Relationship_pos_ter" + Time + ".xlsx";
                strSheetExport = "pos_ter";
                CreateRealaPos_Termi(nameFileNew, strSheetExport, dbFull);
                #endregion

                #region Tách file Relationship Terminal - Other
                nameFileNew = Folder_New + "\\Relationship_ter_other" + Time + ".xlsx";
                strSheetExport = "ter_other";
                CreateRealaTermi_Other(nameFileNew, strSheetExport, dbFull);
                #endregion

                #region Tách file Relationship Kho - Other
                nameFileNew = Folder_New + "\\Relationship_Kho_other_" + Time + ".xlsx";
                strSheetExport = "Kho_other";
                CreateRealaKho_Other(nameFileNew, strSheetExport, dbFull);//(nameFileNew, strSheetExport, dbFull);
                #endregion

                MessageBox.Show("Tách file dữ liệu Relationship thành công. Kiểm tra dữ liệu đã tách tại thư mục: \n" + Folder_New);
                WriteOLog("Tách file dữ liệu Relationship thành công. Kiểm tra dữ liệu đã tách tại thư mục: \n" + Folder_New);

            }
            catch (Exception)
            {
                MessageBox.Show("Có lỗi trong quá trình tách dữ liệu Relationship. Kiểm tra lại log hệ thống");
                WriteOLog("Có lỗi trong quá trình tách dữ liệu Relationship. Kiểm tra lại log hệ thống" + Folder_New);
                return;
                throw;
            }
        }

        private void button1_Click_2(object sender, EventArgs e)
        {
            string Folder_New = string.Empty;
            string nameFileNew = string.Empty;
            string strSheetExport = string.Empty;
            string sheetImport = string.Empty;
            string Time = DateTime.Now.ToLongTimeString().Replace(":", "_").Replace(" ", "_");
            DataTable dbFull, dbBGT, dbSim3G;

            Folder_New = txtFolderLog.Text + "\\" + Time;
            if (!System.IO.Directory.Exists(Folder_New))
            {
                System.IO.Directory.CreateDirectory(Folder_New);
            }

            if (!CheckFileTemplate(txtFile.Text))
            {
                MessageBox.Show("Kiểm tra lại quy cách đặt tên sheet theo đúng mẫu quy định.");
                WriteOLog("Kiểm tra lại quy cách đặt tên sheet theo đúng mẫu quy định.");
                return;
            }

            try
            {
                // Đọc dữ liệu từ sheet FullCI
                sheetImport = "FullCI";
                dbFull = tbReadEx(txtFile.Text, sheetImport);

                // Đọc dữ liệu từ sheet BGT Code
                sheetImport = "BGT";
                dbBGT = tbReadEx(txtFile.Text, sheetImport);

                //Đọc dữ liệu từ sheet Sim3G
                sheetImport = "Sim3G";
                dbSim3G = tbReadEx(txtFile.Text, sheetImport);

                if (!CheckCountColumn(dbFull, dbBGT, dbSim3G))
                {
                    MessageBox.Show("Kiểm tra lại file tách dữ liệu theo đúng mẫu quy định.");
                    WriteOLog("Kiểm tra lại file tách dữ liệu theo đúng mẫu quy định.");
                    return;
                }
            }
            catch (Exception)
            {
                MessageBox.Show("Kiểm tra lại quy cách đặt tên của file và sheet");
                WriteOLog("Kiểm tra lại quy cách đặt tên của file và sheet");
                return;
                throw;
            }

            try
            {
               
                #region Tách file Relationship Kho - Other
                nameFileNew = Folder_New + "\\Relationship_Kho_other_" + Time + ".xlsx";
                strSheetExport = "Kho_other";
                CreatePOS(nameFileNew, strSheetExport, dbFull);//(nameFileNew, strSheetExport, dbFull);
                #endregion

                MessageBox.Show("Tách file dữ liệu Relationship thành công. Kiểm tra dữ liệu đã tách tại thư mục: \n" + Folder_New);
                WriteOLog("Tách file dữ liệu Relationship thành công. Kiểm tra dữ liệu đã tách tại thư mục: \n" + Folder_New);

            }
            catch (Exception)
            {
                MessageBox.Show("Có lỗi trong quá trình tách dữ liệu Relationship. Kiểm tra lại log hệ thống");
                WriteOLog("Có lỗi trong quá trình tách dữ liệu Relationship. Kiểm tra lại log hệ thống" + Folder_New);
                return;
                throw;
            }
        }

       

    }
}
