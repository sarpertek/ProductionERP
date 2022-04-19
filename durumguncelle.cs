using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using MetroFramework.Forms;

namespace _2015
{
    public partial class durumguncelle : MetroForm
    {
        public durumguncelle()
        {
            InitializeComponent();
        }

        private void durumguncelle_Load(object sender, EventArgs e)
        {
        }

        private void metroComboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (metroComboBox1.Text == "Tedarikçi bilgisi")
            {

            }          
            else if (metroComboBox1.Text == "Galvanize gönderildi")
            {

            }
            else if (metroComboBox1.Text == "Galvenizden geldi")
            {

            }
            else if (metroComboBox1.Text == "Katofereze gönderildi")
            {

            }
            else if (metroComboBox1.Text == "Katoferezden geldi")
            {

            }
            else if (metroComboBox1.Text == "GKK'de onaylandı")
            {

            }
            else if (metroComboBox1.Text == "GKK'de reddedildi ")
            {

            }
            else { }
        }
    }
}
