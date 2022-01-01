using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Word;

namespace WordAutoDesktop
{
    public partial class Form1 : Form
    {
        private OpenFileDialog ofd;

        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            ofd = new OpenFileDialog();
        }

        private void tableLayoutPanel1_Paint(object sender, PaintEventArgs e)
        {

        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                label4.Text = ofd.FileName;
                var helper = new WordHelper(ofd.FileName);

                var items = new Dictionary<string, string>
                {
                    { "<TAG1>", textBox1.Text },
                    { "<TAG2>", textBox2.Text },
                };

                helper.Process(items);
            }
        }

        private void label4_Click(object sender, EventArgs e)
        {

        }
    }
}
