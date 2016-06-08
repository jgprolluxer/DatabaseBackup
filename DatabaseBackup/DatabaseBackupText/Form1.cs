using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using DatabaseBackup;

namespace DatabaseBackupText
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //string path = @"C:\Program Files\MySQL\MySQL Server 5.5\bin\mysqldump.exe -u " + txtBoxDBUsername.Text + @" -p " + txtBoxDBName.Text + @" > " + txtBoxDBName.Text + @".sql";
            var args = new[]
            {
                textBox1.Text,
                textBox2.Text,
                textBox3.Text,
                textBox4.Text
            };

            Backup.Database(args);
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            textBox1.Text = @"C:\Program Files\MySQL\MySQL Server 5.5\bin\mysqldump.exe -u ";
            textBox2.Text = @"username";
            textBox3.Text = @"database";
            textBox4.Text = @"filename";
        }
    }
}
