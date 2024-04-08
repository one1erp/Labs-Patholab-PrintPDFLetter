using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace PrintPDFLetter
{
    public partial class Fax_Prompt : Form
    {
        
            private string _phoneNumber="-1";
            public string PhoneNumber
            {
                get { return _phoneNumber; }
                set { _phoneNumber = value; }
            }

          

        public Fax_Prompt()
        {
            InitializeComponent();
        }

        private void textFaxNumber_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                btnAut_Click(null, null);
            }
        }

        private void Fax_Prompt_Load(object sender, EventArgs e)
        {

        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void btnAut_Click(object sender, EventArgs e)
        {
            if(textFaxNumber.Text.Length<7 )
            {
                MessageBox.Show("מספר הטלפון קצר מדי, פחות מ-7 תווים");
                return;

            }
            if (textFaxNumber.Text.Length>11 )
            {
                MessageBox.Show("מספר הטלפון ארוך מדי - מעל 11 תווים");
                return;

            }
            long i;
            if (!long.TryParse(textFaxNumber.Text,out i))
            {
                MessageBox.Show("מספר יכול להכיל ספרות בלבד");
                return;

            }
            _phoneNumber = textFaxNumber.Text;
            this.Close();
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            _phoneNumber = "-1";
            this.Close();
        }
    }
}
