using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace EnergyHackProject
{
    public partial class Form1 : Form
    {

        List<PowerLinesClass> CurListPowerLines;

        List<SubstationClass> CurListSubstation;

        /*<SQL DATA>*/
        const string SqlUser = "energy";
        const string SqlPass = "storage";
        const string SqlHost = "185.21.142.28";
        const string SqlBase = "energy";
        const string SqlPort = "33061";

        MySql.Data.MySqlClient.MySqlConnection mySqlConnection = new MySqlConnection("server=" + SqlHost + ";user id=" + SqlUser + ";database=" + SqlBase + ";port=" + SqlPort + ";password=" + SqlPass + ";");
        /*</SQL DATA>*/

        public Form1()
        {
            InitializeComponent();

            CurListPowerLines = new List<PowerLinesClass>();
            CurListSubstation = new List<SubstationClass>();
        }

        /// <summary>
        /// Импорт ЛЭП
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button1_LEP_ImportFile_Click(object sender, EventArgs e)
        {
            LoaderExcel f = new LoaderExcel(true);
            f.ShowDialog();
            List<PowerLinesClass> buf = f.ListPowerLines;
            if (buf == null) return;



            CurListPowerLines = new List<PowerLinesClass>(buf);
            dataGridView1.DataSource = CurListPowerLines;

            //for (int i = 0; i < dataGridView1.Columns.Count; i++)
            //{
            //    for (int j = 0; j < dataGridView1.Rows.Count; j++)
            //    {
            //        try
            //        {
            //            if (dataGridView1.Rows[j].Cells[i].Value.ToString() == "-") dataGridView1.Rows[j].Cells[i].Value = null;
            //            else if (dataGridView1.Rows[j].Cells[i].Value.ToString() == "-1") dataGridView1.Rows[j].Cells[i].Value = null;
            //        }
            //        catch { }
            //    }
            //}
        }

        /// <summary>
        /// Импорт ПС
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button9_PS_ImportFile_Click(object sender, EventArgs e)
        {
            LoaderExcel f = new LoaderExcel(false);
            f.ShowDialog();
            List<SubstationClass> buf = f.ListSubstation;
            if (buf == null) return;
            CurListSubstation = new List<SubstationClass>(buf);
            dataGridView2.DataSource = CurListSubstation;

            //for (int i = 0; i < dataGridView2.Columns.Count; i++)
            //{
            //    for (int j = 0; j < dataGridView2.Rows.Count; j++)
            //    {
            //        try
            //        {
            //            if (dataGridView2.Rows[j].Cells[i].Value.ToString() == "-") dataGridView2.Rows[j].Cells[i].Value = null;
            //            else if (dataGridView2.Rows[j].Cells[i].Value.ToString() == "-1") dataGridView2.Rows[j].Cells[i].Value = null;
            //        }
            //        catch
            //        {
            //        }
            //    }
            //}
        }

        private void button2_Click(object sender, EventArgs e)
        {

        }

        private async void buttonMySQLConnect_Click(object sender, EventArgs e)
        {
            await connecttoBD();


            //int a = 0;
        }
        Task connecttoBD(bool noMSG = false)
        {
            var dbTask = Task.Run(() => 
                {
                try
                {
                    label_Info_DB.Text = "Подключение к БД..";
                    if (!(mySqlConnection.State == ConnectionState.Open))
                    {
                        mySqlConnection.Open();
                        labelSqlState.Text = "Подключено";
                        label_Info_DB.Text = "Подключено";
                        //buttonMySQLConnect.Text = "Отключиться";
                        button5_LEP_ImportBD.Enabled = true;
                        button4_LEP_ExportBD.Enabled = true;
                        button6_PS_ImportBD.Enabled = true;
                        button7_PS_ExportBD.Enabled = true;
                    }
                    else
                    {
                        button5_LEP_ImportBD.Enabled = false;
                        button4_LEP_ExportBD.Enabled = false;
                        button6_PS_ImportBD.Enabled = false;
                        button7_PS_ExportBD.Enabled = false;
                        mySqlConnection.Close();
                        labelSqlState.Text = "Отключено";
                        label_Info_DB.Text = "Отключено";
                        //buttonMySQLConnect.Text = "Подключиться";
                    }

                }
                catch (Exception ex)
                {
                    if (!noMSG) MessageBox.Show("Error: " + ex.Message.ToString());
                }
            });

            return dbTask;
        }

        public string CorrectCellName(string badCellName, string variTableName)
        {
            try
            {
                //string query = "SELECT variName FROM " + SqlTableVariations + " WHERE id = 1";
                MySqlCommand command = new MySqlCommand("SELECT count(variName) FROM " + variTableName + " WHERE variVariations LIKE '%" + badCellName.Trim() + "%'", mySqlConnection);
                switch (Convert.ToUInt32(command.ExecuteScalar().ToString()))
                {
                    case 1:
                        //Ровно одно совпадение, читаем название колонки и отдаем
                        command = new MySqlCommand("SELECT variName FROM " + variTableName + " WHERE variVariations LIKE '%" + badCellName.Trim() + "%'", mySqlConnection);
                        return command.ExecuteScalar().ToString();
                    case 0:
                        //Совпадений нет
                        //Запрос ручного ввода, отдаем NULL
                        break;
                    default:
                        //MessageBox.Show("err#t79qbk5qg0 - Слишком много совпадений в базе обозначений колонок! Слишком маленькое имя в колонке или в базе обнаружено неоднозначное соответствие имен!");
                        return "ambiguity";
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message.ToString());
            }
            return null;
        }

        public string[] getVariationsNames(string variTableName)
        {
            MySqlCommand command = new MySqlCommand("SELECT variName FROM " + variTableName, mySqlConnection);
            using (var reader = command.ExecuteReader())
            {
                string[] outArr = null;
                //ClassName[] allRecords = null;
                var list = new List<string>();
                while (reader.Read())
                    list.Add(reader.GetString(0));
                outArr = list.ToArray();
                return outArr;
            }
        }

        private void CorrectCellNameButton_Click(object sender, EventArgs e)
        {
            //TotalLength = new CableClass.TotalLength(2, new CableClass.TotalLength
            //PowerLinesClass myPowerLine = new PowerLinesClass("АО РИМ", "Ветка1", "69", "Какойто номерлинии или чо", "Имялинии", 110, 2007,new CableClass(2, new CableClass.TotalLength { LenghtPerСircuitALL = 2, trackLengthALL = 120 }, new CableClass.СircuitLength { LenghtPerСircuit = 3, trackLength=43}, "WTF"),"Норм");
            //PowerLinesInsertRow(myPowerLine);

            //var test = getVariationsNames("variRow");
            //MessageBox.Show(test[0].ToString()+ test[1].ToString());


            //SubstationClass mySubstation = new SubstationClass("АО НИЭП","Branch", "numberPP", "NameRESwtf??","Подстанция какая-то", "Классволтаге", 2007,"Transnumber", "Transtype", 0.542,0.88, 2005, 2008, 0.2, "Good");
            //SubstationsInsertRow(mySubstation);
            string testCommentLine = "1,";


        }

        public bool PowerLinesInsertRow(PowerLinesClass myPowerLine)
        {
            try
            {
                Random random = new Random();
                string sql = "INSERT INTO powerlines (Organization,Branch,NumberPP,NumberPowerLines,NamePowerLines,Voltage,CommissioningYear,CountСircuit,trackLengthALL,LenghtPerСircuitALL,trackLength,LenghtPerСircuit,Model,TechnicalCondition,Comments,LoadedBy,LoadTimestamp,UniqueId) VALUES('" +
                myPowerLine.Organization + "','" +
                myPowerLine.Branch + "','" +
                myPowerLine.NumberPP + "','" +
                myPowerLine.NumberPowerLines + "','" +
                myPowerLine.NamePowerLines + "','" +
                myPowerLine.Voltage + "','" +
                myPowerLine.CommissioningYear + "','" +
                myPowerLine.Cable.CountСircuit + "','" +
                myPowerLine.Cable.totalLength.trackLengthALL + "','" +
                myPowerLine.Cable.totalLength.LenghtPerСircuitALL + "','" +
                myPowerLine.Cable.circuitLength.trackLength + "','" +
                myPowerLine.Cable.circuitLength.LenghtPerСircuit + "','" +
                myPowerLine.Cable.Model + "','" +
                myPowerLine.TechnicalCondition + "','" +
                myPowerLine.Comments + "','" +
                System.Security.Principal.WindowsIdentity.GetCurrent().Name + "','" +
                DateTime.Now.ToString("HH:mm:ss dd-MM-yyyy") + "','" +
                random.Next(00000, 99999) + "')";
                MySqlCommand command = new MySqlCommand(sql, mySqlConnection);
                command.ExecuteNonQuery();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message.ToString());
            }
            return true;
        }

        /*public bool SubstationsGetRow(SubstationClass mySubstation)
        {
            SubstationClass mySubstation = new SubstationClass("АО НИЭП", "Branch", "numberPP", "NameRESwtf??", "Подстанция какая-то", "Классволтаге", 2007, "Transnumber", "Transtype", 0.542, 0.88, 2005, 2008, 0.2, "Good");
            List<SubstationClass> ret = new List<SubstationClass>();
            ret.Add();

            return;
        }
        */


        public bool SubstationsInsertRow(SubstationClass mySubstation)
        {
            try
            {
                Random random = new Random();
                string sql = "INSERT INTO substations (Organization, Branch, NumberPP, NameRES, NameSubstation, ClassVoltage, YearInput, TransNumber, TransType, TransNomPower, TransLoadWinter, TransYearManufacture, TransYearOn, TransPercentWear, TransCondition, LoadedBy, LoadTimestamp, UniqueId) VALUES('" +
                mySubstation.Organization + "','" +
                mySubstation.Branch + "','" +
                mySubstation.NumberPP + "','" +
                mySubstation.NameRES + "','" +
                mySubstation.NameSubstation + "','" +
                mySubstation.ClassVoltage + "','" +
                mySubstation.YearInput + "','" +
                mySubstation.TransNumber + "','" +
                mySubstation.TransType + "','" +
                mySubstation.TransNomPower + "','" +
                mySubstation.TransLoadWinter + "','" +
                mySubstation.TransYearManufacture + "','" +
                mySubstation.TransYearOn + "','" +
                mySubstation.TransPercentWear + "','" +
                mySubstation.TransCondition + "','" +
                System.Security.Principal.WindowsIdentity.GetCurrent().Name + "','" +
                DateTime.Now.ToString("HH:mm:ss dd-MM-yyyy") + "','" +
                random.Next(00000, 99999) + "')";
                MySqlCommand command = new MySqlCommand(sql, mySqlConnection);
                command.ExecuteNonQuery();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message.ToString());
                return false;
            }
            return true;
        }

        private void textBoxSearch_TextChanged(object sender, EventArgs e)
        {
            string result = CorrectCellName(textBoxSearch.Text, "variColumn");
            switch (result)
            {
                case null:
                    labelSearchResult.Text = "Результатов нет";
                    break;
                case "ambiguity":
                    labelSearchResult.Text = "Больше 1го результата!";
                    break;
                default:
                    labelSearchResult.Text = result;
                    break;
            }
        }

        public void button3__LEP_ExportFile_Click_1(object sender, EventArgs e)
        {
            WordCreate wc = new WordCreate();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            ExportToExcel.CreatePLExcel(CurListPowerLines);
        }

        private void button4_LEP_ExportBD_Click(object sender, EventArgs e)
        {
            foreach (var item in CurListPowerLines)
            {
                PowerLinesInsertRow(item);
            }
        }

        private void button8_PS_ExportFile_Click(object sender, EventArgs e)
        {
            ExportToExcel.CreateSSExcel(CurListSubstation);
        }

        private void dataGridView1_CellMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right)
            {
                contextMenuStrip1.Show(MousePosition, ToolStripDropDownDirection.Right);
            }
        }


        private void button5_LEP_ImportBD_Click(object sender, EventArgs e)
        {

        }


        private void button6_PS_ImportBD_Click(object sender, EventArgs e)
        {

        }

        private void tabPage1_Click(object sender, EventArgs e)
        {

        }

        private void contextMenuStrip1_Opening(object sender, CancelEventArgs e)
        {

        }

        private void подключитьсяToolStripMenuItem_Click(object sender, EventArgs e)
        {
            connecttoBD();
        }

        private void button7_PS_ExportBD_Click(object sender, EventArgs e)
        {

        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }
    }
}
