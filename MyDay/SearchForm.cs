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
    public partial class SearchForm : Form
    {
        public SearchForm()
        {
            InitializeComponent();
        }

        private void textBox_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == Convert.ToChar(Keys.Return))
            {
                //MessageBox.Show("Key pressed");
                DialogResult = DialogResult.OK;
            }
        }

        public string GetText()
        {
            return textBox.Text;
        }

        public bool MatchCase()
        {
            return cbMatchCase.Checked;
        }

        public bool WholeWord()
        {
            return cbWholeWord.Checked;
        }

        public bool RegularExpression()
        {
            return cbRegular.Checked;
        }

        private void SearchForm_Activated(object sender, EventArgs e)
        {
            textBox.Focus();
        }

        private void cbRegular_CheckedChanged(object sender, EventArgs e)
        {
            if (cbRegular.Checked)
                cbWholeWord.Checked = false;
        }

        private void cbWholeWord_CheckedChanged(object sender, EventArgs e)
        {
            if (cbWholeWord.Checked)
                cbRegular.Checked = false;
        }
    }
}
