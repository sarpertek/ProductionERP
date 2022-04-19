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
    public partial class Duzenle : MetroForm
    {
        private string isemri, ucisemri, musteriadi, giristarihi, paketlemetarihi, projeadi, durum, sevktarihi, notlar;
        public string passvalue1
        {
            get { return isemri; }
            set { isemri = value; }
        }
        public string passvalue2
        {
            get { return ucisemri; }
            set { ucisemri = value; }
        }
        public string passvalue3
        {
            get { return musteriadi; }
            set { musteriadi = value; }
        }

        public string passvalue4
        {
            get { return giristarihi; }
            set { giristarihi = value; }
        }
        public string passvalue5
        {
            get { return paketlemetarihi; }
            set { paketlemetarihi = value; }
        }
        public string passvalue6
        {
            get { return projeadi; }
            set { projeadi = value; }
        }
        public string passvalue7
        {
            get { return durum; }
            set { durum = value; }
        }
        public string passvalue8
        {
            get { return sevktarihi; }
            set { sevktarihi = value; }
        }
        public string passvalue9
        {
            get { return notlar; }
            set { notlar = value; }
        }
        public Duzenle()
        {
            InitializeComponent();
        }

        private void Duzenle_Load(object sender, EventArgs e)
        {
            metroTextBox1.Text = isemri;
            metroTextBox2.Text = ucisemri;
            metroTextBox3.Text = musteriadi;
            metroTextBox4.Text = giristarihi;
            metroTextBox5.Text = paketlemetarihi;
            metroTextBox6.Text = projeadi;
            metroTextBox7.Text = durum;
            metroTextBox8.Text = sevktarihi;
            label18.Text = notlar;
            label18.ScrollBars = ScrollBars.Vertical;
        }
    }
}
