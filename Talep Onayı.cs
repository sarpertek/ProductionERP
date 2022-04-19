using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using MetroFramework;
using MetroFramework.Forms;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.IO;
using Microsoft.Office.Core;
using System.Data.SQLite;

namespace _2015
{
    public partial class Talep_Onayı : MetroForm
    {
        string toplam, iscilik, kaplama, malzeme;
        private string fasonparca, fasonparcaadedi, fasonparca2, fasonparca3, fasonparca4;
        public string fasonparcastring
        {
            get { return fasonparca; }
            set { fasonparca = value; }
        }
        public string fasonparcaadedistring
        {
            get { return fasonparcaadedi; }
            set { fasonparcaadedi = value; }
        }
        public string fasonparcastring2
        {
            get { return fasonparca2; }
            set { fasonparca2 = value; }
        }
        public string fasonparcastring3
        {
            get { return fasonparca3; }
            set { fasonparca3 = value; }
        }
        public string fasonparcaid
        {
            get { return fasonparca4; }
            set { fasonparca4 = value; }
        }


        public Talep_Onayı()
        {
            InitializeComponent();
        }

        private void metroButton1_Click(object sender, EventArgs e)
        {
            string bugun = DateTime.Now.ToString("dd/MM/yyyy");
            string termin = dateTimePicker1.Value.ToString("dd/MM/yyyy");
            iscilik = "Hariç";
            malzeme = "Hariç";
            kaplama = "Hariç";
            if (iscilikcheck.Checked == true) { iscilik = "Dahil"; }
            if (malzemecheck.Checked == true) { malzeme = "Dahil"; }
            if (kaplamacheck.Checked == true) { kaplama = "Dahil"; }


            string valueA = metroTextBox1.Text.ToString().Replace('.', ',');
            string valueB = metroTextBox3.Text.ToString().Replace('.', ',');
            float result;
            if (float.TryParse(valueA, out result)
                && float.TryParse(valueB, out result))
            {
                float valueC = float.Parse(valueA) * float.Parse(valueB);
                toplam = Math.Round(valueC, 2, MidpointRounding.AwayFromZero).ToString();
            }


            SQLiteConnection baglan = new SQLiteConnection(baglanti.database);
            baglan.Open();

            string sqlfason = "INSERT INTO fasontakibi([Parça no], [Grup adı], [Parça adı], [Fason], [Sipariş tarihi], [Termin], [Sipariş adedi], [Birim fiyat], [Toplam], [Onaylayan], [İşçilik], [Kaplama], [Malzeme]) Values('" + fasonparca + "', '" + fasonparca2 + "', '" + fasonparca3 + "','" + metroTextBox2.Text.ToString() + "', '" + bugun + "', '" + termin + "', '" + metroTextBox1.Text.ToString() + "', '" + metroTextBox3.Text.ToString() + "', '" + toplam + "' , '" + Environment.UserName + "' , '" + iscilik+"','"+kaplama+"','"+malzeme+"' ) ";
            SQLiteCommand cmdfason = new SQLiteCommand(sqlfason, baglan);
            cmdfason.ExecuteNonQuery();
            string sqlfasonbom = "UPDATE bomlistesi SET [Fason] = '" + metroTextBox2.Text.ToString()+ "' , [Durum] = 'Tedarikçide', [Fason Durum] = 'Devam Ediyor', [Op Sayısı] = [Op Sayısı] + 1, [Birim Fiyat] = '"+metroTextBox3.Text.ToString()+"', [Toplam] = '"+toplam+"' WHERE [ID] = '" + fasonparca4+"' ";
            SQLiteCommand cmdfasonbom = new SQLiteCommand(sqlfasonbom, baglan);
            cmdfasonbom.ExecuteNonQuery();
            //KISMİ
            string sqlfasonbom1 = "UPDATE bomlistesi SET [Talaşlı Durum] = 'Devam Ediyor' WHERE [ID] = '" + fasonparca4 + "' AND [Fason Miktarı] = 'Kısmi' AND [Talaşlı Durum] = 'Fason Talep Ediyor'";
            SQLiteCommand cmdfasonbom1 = new SQLiteCommand(sqlfasonbom1, baglan);
            cmdfasonbom1.ExecuteNonQuery();
            string sqlfasonbom2 = "UPDATE bomlistesi SET [Kaynak Durum] = 'Devam Ediyor' WHERE [ID] = '" + fasonparca4 + "' AND [Fason Miktarı] = 'Kısmi' AND [Kaynak Durum] = 'Fason Talep Ediyor' ";
            SQLiteCommand cmdfasonbom2 = new SQLiteCommand(sqlfasonbom2, baglan);
            cmdfasonbom2.ExecuteNonQuery();
            //TAMAMI
            string sqlfasonbom3 = "UPDATE bomlistesi SET [Talaşlı Durum] = 'Bitti' , [Tam Op] = [Tam Op] + 1 WHERE [ID] = '" + fasonparca4 + "' AND [Fason Miktarı] = 'Tamamı' AND [Talaşlı Durum] = 'Fason Talep Ediyor'";
            SQLiteCommand cmdfasonbom3 = new SQLiteCommand(sqlfasonbom3, baglan);
            cmdfasonbom3.ExecuteNonQuery();
            string sqlfasonbom4 = "UPDATE bomlistesi SET [Kaynak Durum] = 'Bitti', [Tam Op] = [Tam Op] + 1 WHERE [ID] = '" + fasonparca4 + "' AND [Fason Miktarı] = 'Tamamı' AND [Kaynak Durum] = 'Fason Talep Ediyor' ";
            SQLiteCommand cmdfasonbom4 = new SQLiteCommand(sqlfasonbom4, baglan);
            cmdfasonbom4.ExecuteNonQuery();


            string sqlfasonlog = "INSERT INTO parcalog([Parça no], [Tarih], [Parça Adedi], [Güncelleyen], [Durum], [Detay]) Values('" + fasonparca + "', '" + bugun + "', '" + metroTextBox1.Text.ToString() + "', '" + Environment.UserName + "' , 'Tedarikçide' , '"+metroTextBox2.Text.ToString()+"' ) ";
            SQLiteCommand cmdfasonlog = new SQLiteCommand(sqlfasonlog, baglan);
            cmdfasonlog.ExecuteNonQuery();

            baglan.Close();
            this.Close();
        }
        
        private void Talep_Onayı_Load(object sender, EventArgs e)
        {
            dateTimePicker1.MinDate = DateTime.Now;
            metroLabel1.Text = fasonparca + " numaralı";
            metroLabel2.Text = fasonparca2 + " " + fasonparca3 + " hangi imalatçıya verildi?";
            metroTextBox1.Text = fasonparcaadedi;
        }
    }
}
