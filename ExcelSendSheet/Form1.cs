using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace ExcelSendSheet
{
    public partial class Form1 : Form
    {
        private Form1()
        {
            InitializeComponent();
        }
        private static Form1 Instance;
        public static void ShowForm()
        {
            if (Instance == null)
            {
                Instance = new Form1();
            }

            Instance.StartPosition = FormStartPosition.CenterParent;
            Instance.ShowDialog();
        }
        private void button1_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
