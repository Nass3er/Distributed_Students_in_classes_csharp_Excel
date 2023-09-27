using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace DistributedStudents
{
    public partial class MainForm : Form
    {
        public MainForm()
        {
            InitializeComponent();
            

        }

        private void button1_Click(object sender, EventArgs e)
        {
            button1.BackColor = Color.White;
            button1.ForeColor = Color.Black;
 
            Form1 frm = new Form1();
            frm.ShowDialog();
            
            button1.BackColor = Color.Navy;
            button1.ForeColor = Color.White;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            button2.BackColor = Color.White;
            button2.ForeColor = Color.Black;
            
            Form2 form2 = new Form2();
            form2.ShowDialog();
             
            button2.BackColor = Color.Navy;
            button2.ForeColor = Color.White;
        }

        private void button3_Click(object sender, EventArgs e)
        {
            button3.BackColor = Color.White;
            button3.ForeColor = Color.Black;
             
            FinalForm final=new FinalForm();
            final.ShowDialog();
             
            button3.BackColor = Color.Navy;
            button3.ForeColor = Color.White;
        }

        private void button4_Click(object sender, EventArgs e)
        {
            button4.BackColor = Color.White;
            button4.ForeColor = Color.Black;
             
            InitialSystem initialSystem = new InitialSystem();
            initialSystem.ShowDialog();
            
            button4.BackColor = Color.Navy;
            button4.ForeColor = Color.White;
        }

        private void MainForm_Load(object sender, EventArgs e)
        {
            //MessageBox.Show(Properties.Settings.Default.finalFilePath);
        }

        private void label8_Click(object sender, EventArgs e)
        {

        }

        private void button6_Click(object sender, EventArgs e)
        {

         }
    }
}
