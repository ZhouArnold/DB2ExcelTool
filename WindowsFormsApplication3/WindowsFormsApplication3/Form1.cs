using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;


namespace WindowsFormsApplication3
{
    public partial class Form1 : Form
    {
        public string datapath = "D:\\Metrics\\Data Source\\";
        public DBAccess _dba = null;
        public ExcelBus _excel = null;
        public Form1()
        {
            InitializeComponent();
            _dba = new DBAccess();
            _excel = new ExcelBus();
        }
        public string SelectFile(string path) {
            Stream myStream = null;
            OpenFileDialog open = new OpenFileDialog();
            open.InitialDirectory = path;
            open.Filter = "txt files (*.xls)|*.xls|All files (*.*)|*.*";
            open.RestoreDirectory = true;

            if (open.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    if ((myStream = open.OpenFile()) != null)
                    {
                        using (myStream)
                        {
                            return open.FileName.ToString();
                        }
                    }
                    else return string.Empty;
                }
                catch (Exception e)
                {
                    MessageBox.Show("Could not read file from disk reason is" + e.Message);
                    return string.Empty;
                }
            }
            else {
                return string.Empty;
            }
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void button4_Click(object sender, EventArgs e)
        {

        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void flowLayoutPanel1_Paint(object sender, PaintEventArgs e)
        {

        }

        private void btnselectRemedy7_Click(object sender, EventArgs e)
        {
            String today = DateTime.Now.ToString("MMdd");
            textBox1.Text = this.SelectFile(datapath+today);

        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (textBox1.Text == "")
            {
                MessageBox.Show("请选择文件！");
            }
            else
            {
                DataSet dataset = new DataSet();
                dataset = _excel.getDataSetFromExcel(textBox1.Text, _excel.GetFirstSheetName(textBox1.Text));
                _dba.ExecuteCommand("delete remedy7");
                _dba.InsertDataToDB(dataset, "remedy7");
                MessageBox.Show("Successed!");
            }
        }

        private void button4_Click_1(object sender, EventArgs e)
        {
            String today = DateTime.Now.ToString("MMdd");
            textBox2.Text = this.SelectFile(datapath + today);
        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (textBox1.Text == "")
            {
                MessageBox.Show("请选择文件！");
            }
            else
            {
                DataSet remedy7IM = new DataSet();
                remedy7IM = _excel.getDataSetFromExcel(textBox2.Text, _excel.GetFirstSheetName(textBox2.Text));
                _dba.ExecuteCommand("delete remedy7IM");
                _dba.InsertDataToDB(remedy7IM, "remedy7IM");
                MessageBox.Show("Successed!");
            }
        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {

        }

        private void button5_Click(object sender, EventArgs e)
        {
            if (textBox1.Text == "")
            {
                MessageBox.Show("请选择文件！");
            }
            else
            {
                ExcelBus excelBiz = new ExcelBus();
                string excelPath = textBox3.Text;
                DataSet dsOpenTask = excelBiz.getDataSetFromExcel(excelPath, "Page1-1$");
                _dba.ExecuteCommand("truncate table TaskAssignedRaw");
                _dba.InsertTaskData(dsOpenTask);
                MessageBox.Show("Successed!");
            }
        }

        private void textBox3_TextChanged(object sender, EventArgs e)
        {

        }

        private void button6_Click(object sender, EventArgs e)
        {
            String today = DateTime.Now.ToString("MMdd");
            textBox3.Text = this.SelectFile(datapath + today);
        }

        private void button1_Click(object sender, EventArgs e)//process daily data
        {
            string weekday = dateTimePicker1.Value.DayOfWeek.ToString();
            string generatedate = dateTimePicker1.Value.ToString("yyyy-MM-dd");
            if (weekday == "Monday")
            {
                MessageBox.Show(generatedate+"是星期一，请使用Process Monday Data按钮");
            }
            else {
                _dba.ExecuteCommand("exec sp_processRemedyData_Auto '" + generatedate + "'");
                _dba.ExecuteCommand("exec sp_processIMSR_Auto '" + generatedate + "'");
                MessageBox.Show("Process remedy7 & remedy7IM data success!");
            }
            
        }

        private void dateTimePicker1_ValueChanged(object sender, EventArgs e)
        {

        }

        private void button8_Click(object sender, EventArgs e)
        {
            if (dateTimePicker1.Enabled == false)
            {
                dateTimePicker1.Enabled = true;
                button8.Text = "Disable Date Select";
            }
            else {
                dateTimePicker1.Enabled = false;
                button8.Text = "Enable Date Select";
            }
            
        }

        private void button7_Click(object sender, EventArgs e)
        {
            if (dateTimePicker1.Value.DayOfWeek.ToString() == "Monday")
            {
                MessageBox.Show(dateTimePicker1.Value.ToString("yyyy-MM-dd") + "是星期一，请使用Generate MON Reports按钮");
            }
            else
            {
                ReportBusiness rptBiz = new ReportBusiness();
                ExcelBus excelBiz = new ExcelBus();
                DateTime generatedate = dateTimePicker1.Value;
                rptBiz.GenerateDailyRptForSingleDays(generatedate);
                //add for generate service remedy group level report to George        
                rptBiz.GenerateServiceRemedyGroupLevel(generatedate);
                //end add by navy
                excelBiz.PrepareFilesForDailyReport(generatedate);
                //MessageBox.Show("Generate Daily Reports Success!");
            }
        }

        private void button9_Click(object sender, EventArgs e)
        {
            string weekday = dateTimePicker1.Value.DayOfWeek.ToString();
            string generatedate = dateTimePicker1.Value.ToString("yyyy-MM-dd");
            if (weekday == "Monday")
            {
                _dba.ExecuteCommand("exec sp_processRemedyData_Auto '" + generatedate +"'");
                _dba.ExecuteCommand("exec sp_processIMSR_Auto '" + generatedate + "'");
                MessageBox.Show("Process Successed!");
            }
            else {
                MessageBox.Show(generatedate + "不是星期一，请使用Process Daily Data按钮");
            }
        }

        private void button10_Click(object sender, EventArgs e)
        {
            string weekday = dateTimePicker1.Value.DayOfWeek.ToString();
            string generatedateday = dateTimePicker1.Value.ToString("yyyy-MM-dd");
            if (weekday == "Monday")
            {
            ReportBusiness rptBiz = new ReportBusiness();
            ExcelBus excelBiz = new ExcelBus();
            DateTime generatedate = dateTimePicker1.Value;
            rptBiz.GenerateDailyRptForMultiDays(generatedate);
            rptBiz.GenerateServiceRemedyGroupLevel(generatedate);
            excelBiz.PrepareFilesForDailyReport(generatedate);
            }
            else
            {
                MessageBox.Show(generatedateday + "不是星期一，请使用Process Daily Data按钮");
            }
           
        }

        private void flowLayoutPanel1_Paint_1(object sender, PaintEventArgs e)
        {

        }

        private void label4_Click(object sender, EventArgs e)
        {

        }

        private void rectangleShape1_Click(object sender, EventArgs e)
        {

        }

        private void button11_Click(object sender, EventArgs e)
        {
            _dba.ExecuteCommand("exec sp_backup_before_OMM ");
            MessageBox.Show("Backup Successed!");
        }

        private void button12_Click(object sender, EventArgs e)
        {
            _dba.ExecuteCommand("exec getOMMbacklog ");
            _dba.ExecuteCommand("exec sp_OMM_recovery");
            MessageBox.Show("Recovery Successed!");
        }

        private void button13_Click(object sender, EventArgs e)
        {
            string weekday = dateTimePicker1.Value.DayOfWeek.ToString();
            if (weekday == "Monday")
            {
                DateTime generatedate = dateTimePicker1.Value;
                ExcelBus excelBiz = new ExcelBus();
                excelBiz.PrepareFilesForRADTricareMonday(generatedate);
                ReportBusiness rptBiz = new ReportBusiness();
                rptBiz.GenerateTricareMonday(generatedate);
            }
            else {
                MessageBox.Show("请选择正确日期！");
            }
        }

        private void textBox4_TextChanged(object sender, EventArgs e)
        {

        }

        private void button14_Click(object sender, EventArgs e)
        {
            String today = DateTime.Now.ToString("MMdd");
            textBox3.Text = this.SelectFile(datapath + today);
        }

        private void button15_Click(object sender, EventArgs e)
        {
            string weekday = dateTimePicker1.Value.DayOfWeek.ToString();
            if (weekday == "Monday")
            {
                ExcelBus excelbiz = new ExcelBus();
                string sheetName = excelbiz.GetFirstSheetName(textBox3.Text);
                DataSet dsRemedy7 = excelbiz.getDataSetFromExcel(textBox3.Text, sheetName);
                _dba.ExecuteCommand("truncate table " + "CorpRemedy");
                _dba.InsertDataToDB(dsRemedy7, "CorpRemedy");
                MessageBox.Show("Insert CorpRemdy Successed!");
            }
            else
            {
                MessageBox.Show("请选择正确日期！");
            }
           
        }

        private void button17_Click(object sender, EventArgs e)
        {
            ExcelBus excelBiz = new ExcelBus();
            DateTime CorpTriWeeklygenerateDate = dateTimePicker1.Value;
            excelBiz.PrepareFilesForCorpTricareMonday(CorpTriWeeklygenerateDate);

            ReportBusiness rptBiz = new ReportBusiness();
            rptBiz.GenerateCorpTriMonday(CorpTriWeeklygenerateDate);
        }

        private void button16_Click(object sender, EventArgs e)
        {
            string weekday = dateTimePicker1.Value.DayOfWeek.ToString();
            if (weekday == "Monday")
            {
                string generatedate = dateTimePicker1.Value.ToString("yyyy-MM-dd");
                
                _dba.ExecuteCommand("exec sp_processCorp_Auto "+generatedate);
                MessageBox.Show("Process Corp Data Successed!");
            }
            else
            {
                MessageBox.Show("请选择正确日期！");
            }
        }

        private void button18_Click(object sender, EventArgs e)
        {
            string weekday = dateTimePicker1.Value.DayOfWeek.ToString();
            if (weekday == "Thursday")
            {
                DateTime TriWeeklygenerateDate = dateTimePicker1.Value;
                ExcelBus excelBiz = new ExcelBus();
                excelBiz.PrepareFilesForTricareRpt(TriWeeklygenerateDate);
                MessageBox.Show("Successed!");
            }
            else
            {
                MessageBox.Show("请选择正确日期！");
            }
        }

        private void button19_Click(object sender, EventArgs e)
        {
            string weekday = dateTimePicker1.Value.DayOfWeek.ToString();
            if (weekday == "Thursday")
            {
                string error = string.Empty;
                ReportBusiness rptBiz = new ReportBusiness();
                DateTime TriWeeklygenerateDate = dateTimePicker1.Value;
                
                rptBiz.GenerateTricareWeeklyRpt(TriWeeklygenerateDate, out error);
                MessageBox.Show("Successed!");
              
            }
            else
            {
                MessageBox.Show("请选择正确日期！");
            }
        }

        private void button20_Click(object sender, EventArgs e)
        {
            String today = DateTime.Now.ToString("MMdd");
            textBox5.Text = this.SelectFile(datapath + today);
        }

        private void textBox5_TextChanged(object sender, EventArgs e)
        {

        }

        private void button21_Click(object sender, EventArgs e)
        {
            if (textBox5.Text == "")
            {
                MessageBox.Show("请选择文件！");
            }
            else
            {
                DataSet dataset = new DataSet();
                dataset = _excel.getDataSetFromExcel(textBox5.Text, _excel.GetFirstSheetName(textBox5.Text));
                _dba.ExecuteCommand("delete PBIpriority");
                _dba.InsertDataToDB(dataset, "PBIpriority");
                MessageBox.Show("Successed!");
            }
        }
    }
}
