using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.ComponentModel.DataAnnotations;
using System.Reflection;
using System.Xml.Serialization;
using System.IO;
using Microsoft.Office.Interop.Excel;

namespace EnergyHackProject
{
    public partial class LoaderExcel : Form
    {
        ConvertXLStoDT LoaderXLS = new ConvertXLStoDT();
        System.Data.DataTable CurrDT;
        List<Sheets> listSheets = new List<Sheets>();

        List<Sheets> listCB2 = new List<Sheets>();

        string FilenameExcel;
        string FilenameTemplatesIn;
        TemplatesInputClass TemplatesInput = new TemplatesInputClass();


        public bool itisLEP;
        public List<PowerLinesClass> ListPowerLines { get; private set; }

        public List<SubstationClass> ListSubstation { get; private set; }

        enum LEPNAME
        {
            [Description("Номер п/п")]
            NumberPP,
            [Description("Диспетчерский номер ЛЭП")]
            NumberPowerLines,
            [Description("Наименование ЛЭП")]
            NamePowerLines,
            [Description("Напряжение")]
            Voltage,
            [Description("Год ввода в эксплуатацию")]
            CommissioningYear,

            [Description("{Кабель} Количество цепей")]
            CountСircuit,

            [Description("{Кабель} Длина всего по трассе")]
            trackLengthALL,
            [Description("{Кабель} Длина всего на 1 цепь")]
            LenghtPerСircuitALL,

            [Description("{Кабель} Длина по участкам по трассе")]
            trackLength,
            [Description("{Кабель} Длина по участкам на 1 цепь")]
            LenghtPerСircuit,
            [Description("{Кабель} Марка")]
            Model,

            [Description("Техническое состояние")]
            TechnicalCondition
        }


        enum PSNAME
        {
            [Description("№ п.п")]
            NumberPP,
            [Description("Наименование РЭС")]
            NameRES,
            [Description("{Подстанция} Наименование")]
            NameSubstation,
            [Description("{Подстанция} Класс напряжения")]
            ClassVoltage,
            [Description("{Подстанция} Год ввода")]
            publicintYearInput,

            [Description("{Трансформатор} №")]
            TransNumber,
            [Description("{Трансформатор} Тип")]
            TransType,
            [Description("{Трансформатор} Номинальная мощность")]
            TransNomPower,
            [Description("{Трансформатор} Загрузка (зимний максимум), % ")]
            TransLoadWinter,
            [Description("{Трансформатор} Год изготовления ")]
            TransYearManufacture,
            [Description("{Трансформатор} Год включения ")]
            TransYearOn,
            [Description("{Трансформатор} % износа*  ")]
            TransPercentWear,//%износа*
            [Description("{Трансформатор} Состояние  ")]
            TransCondition,//Состояние(хор.,удов.,взонериска)


        }

        public LoaderExcel(bool itisLEP_)
        {
            InitializeComponent();
            itisLEP = itisLEP_;

            if (itisLEP) this.Text = "Импорт файлов ЛЭП";
            else this.Text = "Импорт файлов ПС";
        }
        //DataTable CurrDT;
        private void button1_Click(object sender, EventArgs e)
        {


        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                Sheets sheet = (Sheets)comboBox1.SelectedItem;
                CurrDT = LoaderXLS.XLStoDTusingInterOp(FilenameExcel, sheet.Id);
                Parser();
                dataGridView1.DataSource = CurrDT;

                for (int i = 0; i < dataGridView1.Columns.Count; i++)
                {
                    dataGridView1.Columns[i].SortMode = DataGridViewColumnSortMode.NotSortable;
                }
            }
            catch { }
        }




        void Parser()
        {
            if (CurrDT == null) return;
            CurrDT = StripEmptyRows(CurrDT);


            test();
        }


        void test()
        {

            CurrDT.Columns[0].ColumnName = "123";
            CurrDT.AcceptChanges();
        }




        private System.Data.DataTable StripEmptyRows(System.Data.DataTable dt)
        {
            List<int> rowIndexesToBeDeleted = new List<int>();
            int indexCount = 0;
            foreach (var row in dt.Rows)
            {
                var r = (DataRow)row;
                int emptyCount = 0;
                int itemArrayCount = r.ItemArray.Length;
                foreach (var i in r.ItemArray) if (string.IsNullOrWhiteSpace(i.ToString())) emptyCount++;

                if (emptyCount == itemArrayCount) rowIndexesToBeDeleted.Add(indexCount);

                indexCount++;
            }

            int count = 0;
            foreach (var i in rowIndexesToBeDeleted)
            {
                dt.Rows.RemoveAt(i - count);
                count++;
            }

            return dt;
        }

        private void dataGridView1_ColumnHeaderMouseDoubleClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            try
            {
                List<Enum> listNameLep = new List<Enum>();
                if (itisLEP)
                {
                    listNameLep.Add(LEPNAME.NumberPP);
                    listNameLep.Add(LEPNAME.NumberPowerLines);
                    listNameLep.Add(LEPNAME.NamePowerLines);
                    listNameLep.Add(LEPNAME.Voltage);
                    listNameLep.Add(LEPNAME.CommissioningYear);
                    listNameLep.Add(LEPNAME.CountСircuit);
                    listNameLep.Add(LEPNAME.trackLengthALL);
                    listNameLep.Add(LEPNAME.LenghtPerСircuitALL);
                    listNameLep.Add(LEPNAME.trackLength);
                    listNameLep.Add(LEPNAME.LenghtPerСircuit);
                    listNameLep.Add(LEPNAME.Model);
                    listNameLep.Add(LEPNAME.TechnicalCondition);
                }
                else
                {
                    listNameLep.Add(PSNAME.NumberPP);
                    listNameLep.Add(PSNAME.NameRES);
                    listNameLep.Add(PSNAME.NameSubstation);
                    listNameLep.Add(PSNAME.ClassVoltage);
                    listNameLep.Add(PSNAME.publicintYearInput);

                    listNameLep.Add(PSNAME.TransNumber);
                    listNameLep.Add(PSNAME.TransType);
                    listNameLep.Add(PSNAME.TransNomPower);
                    listNameLep.Add(PSNAME.TransLoadWinter);
                    listNameLep.Add(PSNAME.TransYearManufacture);
                    listNameLep.Add(PSNAME.TransYearOn);
                    listNameLep.Add(PSNAME.TransPercentWear);
                    listNameLep.Add(PSNAME.TransCondition);

                }




                ChangleNameColumn ChangeNameLEP = new ChangleNameColumn("", listNameLep);
                ChangeNameLEP.ShowDialog();
                Enum buf = ChangeNameLEP.Enumselect;
                if (buf == null)
                    return;
                if (itisLEP)
                {
                    CurrDT.Columns[e.ColumnIndex].ColumnName = GetDescription((LEPNAME)buf);
                }
                else
                {
                    CurrDT.Columns[e.ColumnIndex].ColumnName = GetDescription((PSNAME)buf);
                }
                dataGridView1.Columns[e.ColumnIndex].SortMode = DataGridViewColumnSortMode.NotSortable;
            }
            catch { }
        }


        public static string GetDescription(Enum enumElement)
        {

            Type type = enumElement.GetType();

            MemberInfo[] memInfo = type.GetMember(enumElement.ToString());
            if (memInfo != null && memInfo.Length > 0)
            {
                object[] attrs = memInfo[0].GetCustomAttributes(typeof(DescriptionAttribute), false);
                if (attrs != null && attrs.Length > 0)
                    return ((DescriptionAttribute)attrs[0]).Description;
            }

            return enumElement.ToString();
        }

        private void numericUpDown1_ValueChanged(object sender, EventArgs e)
        {
            try
            {
                dataGridView1.DefaultCellStyle.BackColor = Color.White;
                dataGridView1.DataSource = CurrDT;
                for (int i = 0; i < numericUpDown1.Value + 1; i++)
                {
                    dataGridView1.Rows[i].DefaultCellStyle.BackColor = Color.White;
                }
                for (int i = 0; i < numericUpDown1.Value; i++)
                {
                    dataGridView1.Rows[i].DefaultCellStyle.BackColor = Color.LightGray;
                }
            }
            catch { }

        }

        private void button4_Click(object sender, EventArgs e)
        {
            try
            {
                if (itisLEP) LoaderLEP();
                else LoaderSubstations();
            }
            catch { }
        }

        void LoaderLEP()
        {
            List<PowerLinesClass> PowerLinesLoadList = new List<PowerLinesClass>();
            CableClass Cable;
            //
            string Organization = currTemplate.temlateName;//Организация //класс
            string Branch = "-";//Филиал  //класс
            string NumberPP = "";//№ п.п.
            string NumberPowerLines = "-"; //Диспетчерский номер ЛЭП
            string NamePowerLines = "-";//Наименование ЛЭП
            int Voltage = -1; //Напряжение, кВ -------------Лёха, это вроде должно быть string, так как никаких расчетов с напряжением не будет. Надо просто показать какой класс напряжения и все. 
            int CommissioningYear = -1;
            //cable
            int CountСircuit = -1;

            bool b1, b2, b3, b4, b5, b6, b7, b8, b9, b10, b11, b12;

            double trackLengthALL = -1, LenghtPerСircuitALL = -1, trackLength = -1, LenghtPerСircuit = -1;
            CableClass.TotalLength totalLength = new CableClass.TotalLength { trackLengthALL = trackLengthALL, LenghtPerСircuitALL = LenghtPerСircuitALL };//на 1 цепь }
            CableClass.СircuitLength circuitLength = new CableClass.СircuitLength { trackLength = trackLength, LenghtPerСircuit = LenghtPerСircuit };
            string Model = "-";

            string TechnicalCondition = "-";
            //
            int bufInt = 0;

            for (int i = (int)numericUpDown1.Value; i < CurrDT.Rows.Count; i++)
            {
                b1 = doljnoPomochLEP(i, LEPNAME.NumberPP, ref NumberPP);
                b2 = doljnoPomochLEP(i, LEPNAME.NumberPowerLines, ref NumberPowerLines);
                b3 = doljnoPomochLEP(i, LEPNAME.NamePowerLines, ref NamePowerLines);
                b4 = doljnoPomochLEP(i, LEPNAME.Voltage, ref Voltage);
                b5 = doljnoPomochLEP(i, LEPNAME.CommissioningYear, ref CommissioningYear);
                b6 = doljnoPomochLEP(i, LEPNAME.CountСircuit, ref CountСircuit);
                b7 = doljnoPomochLEP(i, LEPNAME.Model, ref Model);
                b8 = doljnoPomochLEP(i, LEPNAME.TechnicalCondition, ref TechnicalCondition);

                b9 = doljnoPomochLEP(i, LEPNAME.trackLengthALL, ref trackLengthALL);
                b10 = doljnoPomochLEP(i, LEPNAME.LenghtPerСircuitALL, ref LenghtPerСircuitALL);
                b11 = doljnoPomochLEP(i, LEPNAME.trackLength, ref trackLength);
                b12 = doljnoPomochLEP(i, LEPNAME.LenghtPerСircuit, ref LenghtPerСircuit);

                if (b1 && !b2 && !b3 && !b4 && !b5 && !b6 && !b7 && !b8 && !b9 && !b10 && !b11 && !b12)
                {
                    try
                    {
                        int bt = Convert.ToInt32(NumberPP);
                    }
                    catch
                    {
                        Branch = NumberPP;
                        NumberPP = "-";
                        NumberPowerLines = "-";
                        NamePowerLines = "-";
                        Voltage = -1;
                        CommissioningYear = -1;
                        CountСircuit = -1;
                        Model = "-";
                        TechnicalCondition = "-";
                        trackLengthALL = -1;
                        LenghtPerСircuitALL = -1;
                        trackLength = -1;
                        LenghtPerСircuit = -1;
                    }
                }

                totalLength = new CableClass.TotalLength { trackLengthALL = trackLengthALL, LenghtPerСircuitALL = LenghtPerСircuitALL };//на 1 цепь }
                circuitLength = new CableClass.СircuitLength { trackLength = trackLength, LenghtPerСircuit = LenghtPerСircuit };
                Cable = new CableClass(CountСircuit, totalLength, circuitLength, Model);
                PowerLinesLoadList.Add(new PowerLinesClass(Organization, Branch, NumberPP, NumberPowerLines, NamePowerLines, Voltage, CommissioningYear, Cable, TechnicalCondition));

            }
            int sa = 9;

            ListPowerLines = new List<PowerLinesClass>(PowerLinesLoadList);

            this.Close();
        }


        void LoaderSubstations()
        {
            List<SubstationClass> SubstationLoadList = new List<SubstationClass>();
            //
            string Organization;
            try
            {
                Organization = currTemplate.temlateName;//Организация //класс
            }
            catch
            {
                Organization = "-";
            };
            string Branch = "-";//Филиал  //класс

            string NumberPP = "-";
            string NameRES = "-";
            string NameSubstation = "-";
            string ClassVoltage = "-";
            int intYearInput = -1;

            string TransNumber = "-";
            string TransType = "-";
            double TransNomPower = -1;
            double TransLoadWinter = -1;
            int TransYearManufacture = -1;
            int TransYearOn = -1;
            double TransPercentWear = -1;
            string TransCondition = "-";


            bool b1, b2, b3, b4, b5, b6, b7, b8, b9, b10, b11, b12, b13;



            string TechnicalCondition = "-";
            //
            int bufInt = 0;

            for (int i = (int)numericUpDown1.Value; i < CurrDT.Rows.Count; i++)
            {
                b1 = doljnoPomochPS(i, PSNAME.NumberPP, ref NumberPP);
                b2 = doljnoPomochPS(i, PSNAME.NameRES, ref NameRES);
                b3 = doljnoPomochPS(i, PSNAME.NameSubstation, ref NameSubstation);
                b4 = doljnoPomochPS(i, PSNAME.ClassVoltage, ref ClassVoltage);
                b5 = doljnoPomochPS(i, PSNAME.publicintYearInput, ref intYearInput);

                b6 = doljnoPomochPS(i, PSNAME.TransNumber, ref TransNumber);
                b7 = doljnoPomochPS(i, PSNAME.TransType, ref TransType);
                b8 = doljnoPomochPS(i, PSNAME.TransNomPower, ref TransNomPower);
                b9 = doljnoPomochPS(i, PSNAME.TransLoadWinter, ref TransLoadWinter);
                b10 = doljnoPomochPS(i, PSNAME.TransYearManufacture, ref TransYearManufacture);
                b11 = doljnoPomochPS(i, PSNAME.TransYearOn, ref TransYearOn);
                b12 = doljnoPomochPS(i, PSNAME.TransPercentWear, ref TransPercentWear);
                b13 = doljnoPomochPS(i, PSNAME.TransCondition, ref TransCondition);


                if (b1 && !b2 && !b3 && !b4 && !b5 && !b6 && !b7 && !b8 && !b9 && !b10 && !b11 && !b12 && !b13)
                {
                    try
                    {
                        int bt = Convert.ToInt32(NumberPP);
                    }
                    catch
                    {
                        Branch = NumberPP;
                        NumberPP = "-";
                        NameRES = "-";
                        NameSubstation = "-";
                        ClassVoltage = "-";
                        intYearInput = -1;
                        TransNumber = "-";
                        TransType = "-";
                        TransNomPower = -1;
                        TransLoadWinter = -1;
                        TransYearManufacture = -1;
                        TransYearOn = -1;
                        TransPercentWear = -1;
                        TransCondition = "-";
                    }
                }


                SubstationLoadList.Add(new SubstationClass(Organization, Branch, NumberPP, NameRES, NameSubstation, ClassVoltage, intYearInput, TransNumber, TransType, TransNomPower, TransLoadWinter, TransYearManufacture, TransYearOn, TransPercentWear, TransCondition));

            }
            int sa = 9;

            ListSubstation = new List<SubstationClass>(SubstationLoadList);
            this.Close();
        }

        bool doljnoPomochLEP(int i, Enum EnumEnum, ref int rez)
        {
            int buffrez = 0;
            try
            {
                if (CurrDT.Rows[i][GetDescription((LEPNAME)EnumEnum)] == DBNull.Value)
                    return false;
                if (CurrDT.Rows[i][GetDescription((LEPNAME)EnumEnum)] == null)
                    return false;
                buffrez = Convert.ToInt32(CurrDT.Rows[i][GetDescription((LEPNAME)EnumEnum)]);
                rez = buffrez;
                return true;
            }
            catch
            {
                return false;
            }
        }
        bool doljnoPomochLEP(int i, Enum EnumEnum, ref double rez)
        {
            double buffrez = 0;
            try
            {
                if (CurrDT.Rows[i][GetDescription((LEPNAME)EnumEnum)] == DBNull.Value)
                    return false;
                if (CurrDT.Rows[i][GetDescription((LEPNAME)EnumEnum)] == null)
                    return false;
                buffrez = Convert.ToDouble(CurrDT.Rows[i][GetDescription((LEPNAME)EnumEnum)]);
                rez = buffrez;
                return true;
            }
            catch
            {
                return false;
            }
        }
        bool doljnoPomochLEP(int i, Enum EnumEnum, ref string rez)
        {
            string buffrez = "";
            try
            {
                if (CurrDT.Rows[i][GetDescription((LEPNAME)EnumEnum)] == DBNull.Value)
                    return false;
                if (CurrDT.Rows[i][GetDescription((LEPNAME)EnumEnum)] == null)
                    return false;
                buffrez = Convert.ToString(CurrDT.Rows[i][GetDescription((LEPNAME)EnumEnum)]);
                if (buffrez == "") return false;
                rez = buffrez;
                return true;
            }
            catch
            {
                return false;
            }
        }


        bool doljnoPomochPS(int i, Enum EnumEnum, ref int rez)
        {
            int buffrez = 0;
            try
            {
                if (CurrDT.Rows[i][GetDescription((PSNAME)EnumEnum)] == DBNull.Value)
                    return false;
                if (CurrDT.Rows[i][GetDescription((PSNAME)EnumEnum)] == null)
                    return false;
                buffrez = Convert.ToInt32(CurrDT.Rows[i][GetDescription((PSNAME)EnumEnum)]);
                rez = buffrez;
                return true;
            }
            catch
            {
                return false;
            }
        }
        bool doljnoPomochPS(int i, Enum EnumEnum, ref double rez)
        {
            double buffrez = 0;
            try
            {
                if (CurrDT.Rows[i][GetDescription((PSNAME)EnumEnum)] == DBNull.Value)
                    return false;
                if (CurrDT.Rows[i][GetDescription((PSNAME)EnumEnum)] == null)
                    return false;
                buffrez = Convert.ToDouble(CurrDT.Rows[i][GetDescription((PSNAME)EnumEnum)]);
                rez = buffrez;
                return true;
            }
            catch
            {
                return false;
            }
        }
        bool doljnoPomochPS(int i, Enum EnumEnum, ref string rez)
        {
            string buffrez = "";
            try
            {
                if (CurrDT.Rows[i][GetDescription((PSNAME)EnumEnum)] == DBNull.Value)
                    return false;
                if (CurrDT.Rows[i][GetDescription((PSNAME)EnumEnum)] == null)
                    return false;
                buffrez = Convert.ToString(CurrDT.Rows[i][GetDescription((PSNAME)EnumEnum)]);
                if (buffrez == "") return false;
                rez = buffrez;
                return true;
            }
            catch
            {
                return false;
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();

        }

        private void button3_Click(object sender, EventArgs e)
        {

        }





        void saveToFile(string filename)
        {
            try
            {
                XmlSerializer formatter = new XmlSerializer(typeof(TemplatesInputClass));
                using (FileStream fs = new FileStream(filename, FileMode.OpenOrCreate))
                {
                    formatter.Serialize(fs, TemplatesInput);
                }
            }
            catch
            { }
        }

        void loadFromFile(string filename)
        {
            try
            {
                XmlSerializer formatter = new XmlSerializer(typeof(TemplatesInputClass));
                TemplatesInputClass tmp;
                using (FileStream fs = new FileStream(filename, FileMode.OpenOrCreate))
                {
                    tmp = (TemplatesInputClass)formatter.Deserialize(fs);
                }
                TemplatesInput = new TemplatesInputClass(tmp);
            }
            catch
            { }
        }


        void saveToBD()//Шаблоны input
        {
            try
            {

                //TO DO

            }
            catch
            { }
        }


        void loadFromBD()//Шаблоны input
        {
            try
            {

                //TO DO

            }
            catch
            { }
        }

        TemplatesInputClass.template currTemplate;
        static int gsdjgalgjdksa = 0;
        private void button5_Click(object sender, EventArgs e)
        {

        }

        private void comboBox2_SelectedIndexChanged_1(object sender, EventArgs e)
        {
            TemplatesInputClass.template TMP = (TemplatesInputClass.template)comboBox2.SelectedItem;
            if (TMP == null) return;
            currTemplate = new TemplatesInputClass.template(TMP);
            if (CurrDT == null) return;
            foreach (var item in currTemplate.TMP)
            {
                try
                {
                    CurrDT.Columns[item.DTColumn].ColumnName = item.Collumn;
                    dataGridView1.Columns[item.DTColumn].SortMode = DataGridViewColumnSortMode.NotSortable;
                }
                catch { }
            }
            numericUpDown1.Value = currTemplate.countRowIgnore;
        }



        private void LoaderExcel_Load(object sender, EventArgs e)
        {
            try
            {
                loadFromFile("defaultTemplatesInput");

                TemplatesInput = new TemplatesInputClass(TemplatesInput);
                comboBox2.DataSource = TemplatesInput.TemplateList;
                comboBox2.DisplayMember = "temlateName";
                comboBox2.ValueMember = "ID";
            }
            catch
            { }
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void открытьФайлToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (openFileDialog1.ShowDialog() == DialogResult.Cancel)
                return;
            // получаем выбранный файл
            FilenameExcel = openFileDialog1.FileName;
            label1.Text = "Текущий файл: " + openFileDialog1.SafeFileName;

            listSheets = LoaderXLS.XLSgetListSHeets(FilenameExcel);
            comboBox1.DataSource = listSheets;
            comboBox1.DisplayMember = "Name";
            comboBox1.ValueMember = "Id";
        }

        private void открытьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (openFileDialog1.ShowDialog() == DialogResult.Cancel)
                return;
            // получаем выбранный файл
            FilenameTemplatesIn = openFileDialog1.FileName;

            loadFromFile(FilenameTemplatesIn);

            TemplatesInput = new TemplatesInputClass(TemplatesInput);
            comboBox2.DataSource = TemplatesInput.TemplateList;
            comboBox2.DisplayMember = "temlateName";
            comboBox2.ValueMember = "ID";
        }

        private void сохранитьКакToolStripMenuItem_Click(object sender, EventArgs e)
        {
            SaveTemploFile();
        }

        void SaveTemploFile(bool isFileDefault = false)
        {
            foreach (var item in TemplatesInput.TemplateList)
            {
                if (item.ID == currTemplate.ID)
                {
                    List<TemplatesInputClass.tCol> buff = new List<TemplatesInputClass.tCol>();
                    if (CurrDT == null) break;
                    for (int i = 0; i < CurrDT.Columns.Count; i++)
                    {
                        buff.Add(new TemplatesInputClass.tCol { Collumn = CurrDT.Columns[i].ColumnName, DTColumn = i });
                    }

                    item.TMP = new List<TemplatesInputClass.tCol>(buff);
                    item.countRowIgnore = Convert.ToInt32(numericUpDown1.Value);
                    break;
                }
            }
            if (isFileDefault) { FilenameTemplatesIn = "defaultTemplatesInput"; }
            else
            {
                if (saveFileDialog1.ShowDialog() == DialogResult.Cancel)
                    return;
                // получаем выбранный файл
                FilenameTemplatesIn = saveFileDialog1.FileName;
            }
            saveToFile(FilenameTemplatesIn);
        }

        private void шаблоныToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }




        private void добавитьToolStripMenuItem_Click(object sender, EventArgs e)
        {

            setString testDialog = new setString();
            string namebuf = "";
            // Show testDialog as a modal dialog and determine if DialogResult = OK.
            if (testDialog.ShowDialog(this) == DialogResult.OK)
            {
                // Read the contents of testDialog's TextBox.
                namebuf = testDialog.RezSTR;
            }
            else
            {
                return;
            }
            testDialog.Dispose();




            TemplatesInput.TemplateList.Add(new TemplatesInputClass.template(namebuf));
            TemplatesInput = new TemplatesInputClass(TemplatesInput);
            comboBox2.DataSource = TemplatesInput.TemplateList;
            comboBox2.DisplayMember = "temlateName";
            comboBox2.ValueMember = "ID";
        }

        private void изменитьИмяToolStripMenuItem_Click(object sender, EventArgs e)
        {

            setString testDialog = new setString();
            string namebuf = "";
            // Show testDialog as a modal dialog and determine if DialogResult = OK.
            if (testDialog.ShowDialog(this) == DialogResult.OK)
            {
                // Read the contents of testDialog's TextBox.
                namebuf = testDialog.RezSTR;
            }
            else
            {
                return;
            }
            testDialog.Dispose();

            foreach (var item in TemplatesInput.TemplateList)
            {
                if (item.ID == currTemplate.ID)
                {
                    TemplatesInput.TemplateList[item.ID].temlateName = namebuf;
                    break;
                }
            }
            TemplatesInput = new TemplatesInputClass(TemplatesInput);
            comboBox2.DataSource = TemplatesInput.TemplateList;
            comboBox2.DisplayMember = "temlateName";
            comboBox2.ValueMember = "ID";
        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            this.Close();
        }

        private void сохранитьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            SaveTemploFile(true);
        }

        private void label4_Click(object sender, EventArgs e)
        {

        }
    }

    [Serializable]
    public class TemplatesInputClass
    {
        public List<template> TemplateList;
        public static int dsadasdas = 0;
        public TemplatesInputClass()
        {
            TemplateList = new List<template>();
        }

        public TemplatesInputClass(TemplatesInputClass n_)
        {
            TemplateList = new List<template>(n_.TemplateList);
        }

        public class template
        {
            public string temlateName { get; set; }

            public int ID { get; set; }
            public int countRowIgnore { get; set; }
            public List<tCol> TMP { get; set; }

            public template()
            {
                temlateName = "";
                TMP = new List<tCol>();
                countRowIgnore = 0;
                ID = dsadasdas++;
            }
            public template(string name_)
            {
                temlateName = name_;
                TMP = new List<tCol>();
                countRowIgnore = 0;
                ID = dsadasdas++;
            }
            public template(template n_)
            {
                temlateName = n_.temlateName;
                TMP = new List<tCol>(n_.TMP);
                countRowIgnore = n_.countRowIgnore;
                ID = n_.ID;
            }
        }

        public class tCol
        {
            public string Collumn { get; set; }
            public int DTColumn { get; set; }

            public tCol()
            {
                Collumn = null;
                DTColumn = 0;
            }
            public tCol(tCol n_)
            {
                Collumn = n_.Collumn;
                DTColumn = n_.DTColumn;
            }
        }


    }


    class Sheets
    {
        public int Id { get; set; }
        public string Name { get; set; }
    }
    class ConvertXLStoDT
    {
        private StringBuilder errorMessages;
        public StringBuilder ErrorMessages
        {
            get { return errorMessages; }
            set { errorMessages = value; }
        }
        public ConvertXLStoDT()
        {
            ErrorMessages = new StringBuilder();
        }

        public List<Sheets> XLSgetListSHeets(string FilePath)
        {
            List<Sheets> rez = new List<Sheets>();
            Microsoft.Office.Interop.Excel.Application excelApp = null;
            Microsoft.Office.Interop.Excel.Workbook workbook = null;
            System.Data.DataTable dt = new System.Data.DataTable(); //Creating datatable to read the content of the Sheet in File.
            try
            {
                excelApp = new Microsoft.Office.Interop.Excel.Application(); // Initialize a new Excel reader. Must be integrated with an Excel interface object.
                //Opening Excel file(myData.xlsx)
                workbook = excelApp.Workbooks.Open(FilePath, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value);

                for (int i = 0; i < workbook.Sheets.Count; i++)
                {
                    //result[0] = ((Excel.Worksheet)wb.Sheets[i+1]).Name;
                    rez.Add(new Sheets { Id = i + 1, Name = ((Worksheet)workbook.Sheets[i + 1]).Name });
                }
            }
            catch (Exception ex)
            {
                ErrorMessages.Append(ex.Message);
            }
            finally
            {
                #region Clean Up                
                if (workbook != null)
                {
                    #region Clean Up Close the workbook and release all the memory.
                    workbook.Close(false, FilePath, Missing.Value);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);
                    #endregion
                }
                workbook = null;
                if (excelApp != null)
                {
                    excelApp.Quit();
                }
                excelApp = null;
                #endregion
            }
            return rez;
        }

        public System.Data.DataTable XLStoDTusingInterOp(string FilePath, int sheetID)
        {

            Microsoft.Office.Interop.Excel.Application excelApp = null;
            Microsoft.Office.Interop.Excel.Workbook workbook = null;
            System.Data.DataTable dt = new System.Data.DataTable(); //Creating datatable to read the content of the Sheet in File.
            try
            {
                excelApp = new Microsoft.Office.Interop.Excel.Application(); // Initialize a new Excel reader. Must be integrated with an Excel interface object.
                //Opening Excel file(myData.xlsx)
                workbook = excelApp.Workbooks.Open(FilePath, Missing.Value, 1, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value);
                Microsoft.Office.Interop.Excel.Worksheet ws = (Microsoft.Office.Interop.Excel.Worksheet)workbook.Sheets.get_Item(sheetID);
                Microsoft.Office.Interop.Excel.Range excelRange = ws.UsedRange; //gives the used cells in sheet
                ws = null; // now No need of this so should expire.
                //Reading Excel file.               
                object[,] valueArray = (object[,])excelRange.get_Value(Microsoft.Office.Interop.Excel.XlRangeValueDataType.xlRangeValueDefault);
                excelRange = null; // you don't need to do any more Interop. Now No need of this so should expire.
                dt = ProcessObjects(valueArray);





            }
            catch (Exception ex)
            {
                ErrorMessages.Append(ex.Message);
            }
            finally
            {
                #region Clean Up                
                if (workbook != null)
                {
                    #region Clean Up Close the workbook and release all the memory.
                    workbook.Close(false, FilePath, Missing.Value);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);
                    #endregion
                }
                workbook = null;
                if (excelApp != null)
                {
                    excelApp.Quit();
                }
                excelApp = null;
                #endregion
            }
            return (dt);
        }
        /// <summary>
        /// Scan the selected Excel workbook and store the information in the cells
        /// for this workbook in an object[,] array. Then, call another method
        /// to process the data.
        /// </summary>
        private void ExcelScanIntenal(Microsoft.Office.Interop.Excel.Workbook workBookIn)
        {
            //
            // Get sheet Count and store the number of sheets.
            //
            int numSheets = workBookIn.Sheets.Count;
            //
            // Iterate through the sheets. They are indexed starting at 1.
            //
            for (int sheetNum = 1; sheetNum < numSheets + 1; sheetNum++)
            {
                Worksheet sheet = (Worksheet)workBookIn.Sheets[sheetNum];
                //
                // Take the used range of the sheet. Finally, get an object array of all
                // of the cells in the sheet (their values). You can do things with those
                // values. See notes about compatibility.
                //
                Microsoft.Office.Interop.Excel.Range excelRange = sheet.UsedRange;
                object[,] valueArray = (object[,])excelRange.get_Value(XlRangeValueDataType.xlRangeValueDefault);
                //
                // Do something with the data in the array with a custom method.
                //
                ProcessObjects(valueArray);
            }
        }
        private System.Data.DataTable ProcessObjects(object[,] valueArray)
        {
            if (valueArray == null) return null;
            System.Data.DataTable dt = new System.Data.DataTable();
            #region Get the COLUMN names
            for (int k = 1; k <= valueArray.GetLength(1); k++)
            {
                dt.Columns.Add((string)valueArray[1, k]);  //add columns to the data table.
            }
            #endregion
            #region Load Excel SHEET DATA into data table
            object[] singleDValue = new object[valueArray.GetLength(1)];
            //value array first row contains column names. so loop starts from 2 instead of 1
            for (int i = 2; i <= valueArray.GetLength(0); i++)
            {
                for (int j = 0; j < valueArray.GetLength(1); j++)
                {
                    if (valueArray[i, j + 1] != null)
                    {
                        singleDValue[j] = valueArray[i, j + 1].ToString();
                    }
                    else
                    {
                        singleDValue[j] = valueArray[i, j + 1];
                    }
                }
                dt.LoadDataRow(singleDValue, System.Data.LoadOption.PreserveChanges);
            }
            #endregion
            return (dt);
        }
    }
}
