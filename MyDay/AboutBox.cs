using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace MyDay
{
    public partial class AboutBox : Form
    {
        public AboutBox()
        {
            InitializeComponent();
        }

        private void AboutBox_Activated(object sender, EventArgs e)
        {
            DateTime now = DateTime.Now;
            // display an easter egg on this special day
            if (now.Month == 2 && now.Day == 7)
                pictureBox1.Image = Properties.Resources.greatwall; // me at the great wall :-)
            else
                pictureBox1.Image = Properties.Resources.logo1; 
        }
    }
}
