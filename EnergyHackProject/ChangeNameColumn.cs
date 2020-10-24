using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace EnergyHackProject
{
    public partial class ChangleNameColumn : Form
    {
        public Enum Enumselect { get; private set; }
        public Enum EnumselectBUF { get; private set; }
        List<changName> listNewName;
        public class changName
        {
            public Enum Id { get; set; }
            public string Name { get; set; }
        }

        public ChangleNameColumn(string nameLast, List<Enum> inEnumsList)
        {
            InitializeComponent();

            listNewName = new List<changName>();
            foreach (var item in inEnumsList)
            {
                listNewName.Add(new changName { Id = item, Name = LoaderExcel.GetDescription(item) });
            }
            comboBox1.DataSource = listNewName;
            comboBox1.DisplayMember = "Name";
            comboBox1.ValueMember = "Id";
        }

        private void ChangleNameColumn_Load(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            Enumselect = EnumselectBUF;
            this.Close();
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            changName changNamet = (changName)comboBox1.SelectedItem;
            EnumselectBUF = changNamet.Id;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Enumselect = null;
            this.Close();
        }
    }
}
