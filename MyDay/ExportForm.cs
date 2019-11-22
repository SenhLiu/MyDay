using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace MyDay
{
    public partial class ExportForm : Form
    {
        public ExportForm()
        {
            InitializeComponent();
        }

        public void SetFromDate(string s)
        {
            txtFromDate.Text = s;
        }

        public void SetToDate(string s)
        {
            txtToDate.Text = s;
        }

        public string GetFromDate()
        {
            return txtFromDate.Text;
        }

        public string GetToDate()
        {
            return txtToDate.Text;
        }

        private void ExportForm_Activated(object sender, EventArgs e)
        {
            txtFromDate.Focus();
        }
    }
}
