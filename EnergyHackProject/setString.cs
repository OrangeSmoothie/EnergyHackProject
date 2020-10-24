using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace EnergyHackProject
{
    public partial class setString : Form
    {
        public string RezSTR { get; private set; }
        public setString()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            RezSTR = textBox1.Text;
        }
    }
}
