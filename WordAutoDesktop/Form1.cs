using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;

namespace WordAutoDesktop
{
    public partial class Form1 : Form
    {
        private OpenFileDialog mainOfd;
        private OpenFileDialog extraOfd;
        private OpenFileDialog testOfd;

        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            mainOfd = new OpenFileDialog();
            extraOfd = new OpenFileDialog();
            testOfd = new OpenFileDialog();
        }

        private void tableLayoutPanel1_Paint(object sender, PaintEventArgs e)
        {

        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (mainOfd.ShowDialog() == DialogResult.OK)
            {
                label4.Text = mainOfd.FileName;
            }
        }

        private void label4_Click(object sender, EventArgs e)
        {

        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (extraOfd.ShowDialog() == DialogResult.OK)
            {
                label5.Text = extraOfd.FileName;
            }
        }

        private void label5_Click(object sender, EventArgs e)
        {

        }

        private void button3_Click(object sender, EventArgs e)
        {
            var helper = new WordHelper(mainOfd.FileName, extraOfd.FileName);
            var items = new Dictionary<string, string>
                {
                    { "<TAG1>", textBox1.Text },
                    { "<TAG2>", textBox2.Text },
                };

            helper.Process(items);
        }

        private void button4_Click(object sender, EventArgs e)
        {
            var testHelper = new WordHelper();
            testHelper.Test();
        }
    }
}
