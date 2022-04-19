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
    public partial class parcahazir : MetroForm
    {
        string sqldurum, sqlimalatbitti, sql3;
        private string parcaid, bolum, bolum2, bolum3, parcano,parcaadet,parcadurumu;
        public string parcaidsi
        {
            get { return parcaid; }
            set { parcaid = value; }
        }
        public string bolumidsi
        {
            get { return bolum; }
            set { bolum = value; }
        }
        public string bolumidsi2
        {
            get { return bolum2; }
            set { bolum2 = value; }
        }
        public string bolumidsi3
        {
            get { return bolum3; }
            set { bolum3 = value; }
        }
        public string parcanosu
        {
            get { return parcano; }
            set { parcano = value; }
        }
        public string parcaadedi
        {
            get { return parcaadet; }
            set { parcaadet = value; }
        }
        public string parcadurum
        {
            get { return parcadurumu; }
            set { parcadurumu = value; }
        }

        private void parcahazir_Load(object sender, EventArgs e)
        {
            metroTextBox1.Text = parcaadet;
            if (bolum == "Kaynak Durum")
            {
                bolum2 = "Talaşlı Durum";
                bolum3 = "Fason Durum";
                metroComboBox1.Items.Remove("Kaynaklı İmalat");
            }
            else if (bolum == "Talaşlı Durum")
            {
                bolum2 = "Kaynak Durum";
                bolum3 = "Fason Durum";
                metroComboBox1.Items.Remove("Talaşlı İmalat");
            }
            else if (bolum == "GKK Red")
            {
                metroComboBox1.Items.Remove("Galvaniz");
                metroComboBox1.Items.Remove("Kataforez");
                metroComboBox1.Items.Remove("Boya");
                metroComboBox1.Items.Remove("Galvanize Gidecek");
                metroComboBox1.Items.Remove("Kalite Kontrol");
                metroComboBox1.Items.Remove("Montaj");
                metroComboBox1.Items.Remove("Sevkiyat");
            }
        }
        public parcahazir()
        {
            InitializeComponent();
        }
        private void metroButton1_Click(object sender, EventArgs e)
        {
            string anlikadet = metroTextBox1.Text.ToString();
            string usrnme = Environment.UserName;
            string bugun = DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss");
            if (bolum == "Kaynak Durum" || bolum == "Talaşlı Durum")
            {
                SQLiteConnection baglan = new SQLiteConnection(baglanti.database);
                baglan.Open();             

                if (metroComboBox1.Text.ToString() == "")
                {
                    MessageBox.Show("Lütfen bir operasyon seçiniz");
                }
                else if (metroComboBox1.Text.ToString() == "Talaşlı İmalat")
                {
                    //Kaynaktan gelen
                    sqldurum = "UPDATE bomlistesi SET [" + bolum + "] = 'Bitti', [Tam Op] = [Tam Op] + 1 , [Gerçek Adet] = '"+anlikadet+"' WHERE [ID] = '" + parcaid + "' ";
                    sqlimalatbitti = "UPDATE bomlistesi SET [Talaşlı Durum] = 'Devam Ediyor', [Op Sayısı] = [Op Sayısı] + 1 WHERE [ID] = '" + parcaid + "' AND (([" + bolum2 + "] = 'Bitti' OR [" + bolum2 + "] is NULL) AND ([" + bolum3 + "] = 'Bitti' OR [" + bolum3 + "] is NULL)) ";
                    MessageBox.Show("Parça 'Talaşlı İmalatta' olarak işaretlenmiştir.");
                }
                else if (metroComboBox1.Text.ToString() == "Kaynaklı İmalat")
                {
                    //Talaşlıdan gelen
                    sqldurum = "UPDATE bomlistesi SET [" + bolum + "] = 'Bitti', [Tam Op] = [Tam Op] + 1 , [Gerçek Adet] = '" + anlikadet + "' WHERE [ID] = '" + parcaid + "' ";
                    sqlimalatbitti = "UPDATE bomlistesi SET [Kaynak Durum] = 'Devam Ediyor', [Op Sayısı] = [Op Sayısı] + 1 WHERE [ID] = '" + parcaid + "' AND (([" + bolum2 + "] = 'Bitti' OR [" + bolum2 + "] is NULL) AND ([" + bolum3 + "] = 'Bitti' OR [" + bolum3 + "] is NULL)) ";
                    MessageBox.Show("Parça 'Kaynaklı İmalatta' olarak işaretlenmiştir.");
                }
                else if (metroComboBox1.Text.ToString() == "Boya")
                { 
                    sqldurum = "UPDATE bomlistesi SET [" + bolum + "] = 'Bitti', [Tam Op] = [Tam Op] + 1 , [Gerçek Adet] = '" + anlikadet + "' WHERE [ID] = '" + parcaid + "' ";
                    sqlimalatbitti = "UPDATE bomlistesi SET [Durum] = 'Boyanacak', [Op Sayısı] = [Op Sayısı] + 1 WHERE [ID] = '" + parcaid + "' AND (([" + bolum2 + "] = 'Bitti' OR [" + bolum2 + "] is NULL) AND ([" + bolum3 + "] = 'Bitti' OR [" + bolum3 + "] is NULL)) ";
                    MessageBox.Show("Parça 'Boyanacak' olarak işaretlenmiştir.");
                }
                else if (metroComboBox1.Text.ToString() == "Galvaniz")
                {
                    sqldurum = "UPDATE bomlistesi SET [" + bolum + "] = 'Bitti', [Tam Op] = [Tam Op] + 1 , [Gerçek Adet] = '" + anlikadet + "' WHERE [ID] = '" + parcaid + "' ";
                    sqlimalatbitti = "UPDATE bomlistesi SET [Durum] = 'Galvanize Gidecek', [Op Sayısı] = [Op Sayısı] + 1 WHERE [ID] = '" + parcaid + "' AND (([" + bolum2 + "] = 'Bitti' OR [" + bolum2 + "] is NULL) AND ([" + bolum3 + "] = 'Bitti' OR [" + bolum3 + "] is NULL)) ";
                    MessageBox.Show("Parça 'Galvanize Gidecek' olarak işaretlenmiştir.");
                }
                else if (metroComboBox1.Text.ToString() == "Kataforez")
                { 
                    sqldurum = "UPDATE bomlistesi SET [" + bolum + "] = 'Bitti', [Tam Op] = [Tam Op] + 1 , [Gerçek Adet] = '" + anlikadet + "' WHERE [ID] = '" + parcaid + "' ";
                    sqlimalatbitti = "UPDATE bomlistesi SET [Durum] = 'Kataforeze Gidecek', [Op Sayısı] = [Op Sayısı] + 1 WHERE [ID] = '" + parcaid + "' AND (([" + bolum2 + "] = 'Bitti' OR [" + bolum2 + "] is NULL) AND ([" + bolum3 + "] = 'Bitti' OR [" + bolum3 + "] is NULL)) ";
                    MessageBox.Show("Parça 'Kataforeze Gidecek' olarak işaretlenmiştir.");
                }
                else if (metroComboBox1.Text.ToString() == "Kalite Kontrol")
                {
                    sqldurum = "UPDATE bomlistesi SET [" + bolum + "] = 'Bitti', [Tam Op] = [Tam Op] + 1 , [Gerçek Adet] = '" + anlikadet + "' WHERE [ID] = '" + parcaid + "' ";
                    sqlimalatbitti = "UPDATE bomlistesi SET [Durum] = 'GKK Bekleniyor', [Op Sayısı] = [Op Sayısı] + 1 WHERE [ID] = '" + parcaid + "' AND (([" + bolum2 + "] = 'Bitti' OR [" + bolum2 + "] is NULL) AND ([" + bolum3 + "] = 'Bitti' OR [" + bolum3 + "] is NULL)) ";
                    MessageBox.Show("Parça 'GKK Bekleniyor' olarak işaretlenmiştir.");
                }
                else if (metroComboBox1.Text.ToString() == "Fason")
                { 
                    sqldurum = "UPDATE bomlistesi SET [" + bolum + "] = 'Bitti', [Tam Op] = [Tam Op] + 1 , [Gerçek Adet] = '" + anlikadet + "' WHERE [ID] = '" + parcaid + "' ";
                    sqlimalatbitti = "UPDATE bomlistesi SET [Durum] = 'Tedarikçiye Gidecek', [Op Sayısı] = [Op Sayısı] + 1 WHERE [ID] = '" + parcaid + "' AND (([" + bolum2 + "] = 'Bitti' OR [" + bolum2 + "] is NULL) AND ([" + bolum3 + "] = 'Bitti' OR [" + bolum3 + "] is NULL)) ";
                    MessageBox.Show("Parça 'Tedarikçiye Gidecek' olarak işaretlenmiştir.");
                }
                else if (metroComboBox1.Text.ToString() == "Montaj")
                {
                    sqldurum = "UPDATE bomlistesi SET [" + bolum + "] = 'Bitti', [Tam Op] = [Tam Op] + 1 , [Gerçek Adet] = '" + anlikadet + "' WHERE [ID] = '" + parcaid + "' ";
                    sqlimalatbitti = "UPDATE bomlistesi SET [Durum] = 'Montaja Hazır', [Montaj Durum] = 'Devam Ediyor', [Op Sayısı] = [Op Sayısı] + 1 WHERE [ID] = '" + parcaid + "' AND (([" + bolum2 + "] = 'Bitti' OR [" + bolum2 + "] is NULL) AND ([" + bolum3 + "] = 'Bitti' OR [" + bolum3 + "] is NULL)) ";
                    MessageBox.Show("Parça 'Montaja Hazır' olarak işaretlenmiştir.");
                }
                else if (metroComboBox1.Text.ToString() == "Sevkiyat")
                {
                    sqldurum = "UPDATE bomlistesi SET [" + bolum + "] = 'Bitti', [Tam Op] = [Tam Op] + 1 , [Gerçek Adet] = '" + anlikadet + "' WHERE [ID] = '" + parcaid + "' ";
                    sqlimalatbitti = "UPDATE bomlistesi SET [Durum] = 'Sevke Hazır' WHERE [ID] = '" + parcaid + "' AND (([" + bolum2 + "] = 'Bitti' OR [" + bolum2 + "] is NULL) AND ([" + bolum3 + "] = 'Bitti' OR [" + bolum3 + "] is NULL)) ";                 
                    MessageBox.Show("Parça 'Sevke Hazır' olarak işaretlenmiştir.");
                }
                SQLiteCommand cmddurum = new SQLiteCommand(sqldurum, baglan);
                cmddurum.ExecuteNonQuery();
                SQLiteCommand cmdimalatbitti = new SQLiteCommand(sqlimalatbitti, baglan);
                cmdimalatbitti.ExecuteNonQuery();
                if (parcadurum == "Tedarikçide")
                {
                    string sqlfasonbitti = "UPDATE bomlistesi SET [Fason Durum] = 'Bitti' WHERE [ID] = '" + parcaid + "' ";
                    SQLiteCommand cmdfasonbitti = new SQLiteCommand(sqlfasonbitti, baglan);
                    cmdfasonbitti.ExecuteNonQuery();
                }
                baglan.Close();
                this.Close();
            }
            else if (bolum == "GKK Onay")
            {
                SQLiteConnection baglan = new SQLiteConnection(baglanti.database);
                baglan.Open();
                sql3 = "INSERT OR IGNORE INTO parcalog([Parça No], [Parça Adedi], [Durum], [Güncelleyen], [Tarih]) VALUES('" + parcanosu + "', '" + parcaadedi + "', '" + bolum + "', '" + usrnme + "', '" + bugun + "') ";
                if (metroComboBox1.Text.ToString() == "")
                {
                    MessageBox.Show("Lütfen bir operasyon seçiniz");
                }
                else if (metroComboBox1.Text.ToString() == "Talaşlı İmalat")
                {
                    //Talaşlıdan gelen
                    sqldurum = "UPDATE bomlistesi SET [Durum] = '" + bolum + "' , [Tam Op] = [Tam Op] + 1 , [Gerçek Adet] = '" + anlikadet + "' WHERE [ID] = '" + parcaid + "' ";
                    sqlimalatbitti = "UPDATE bomlistesi SET [Talaşlı Durum] = 'Devam Ediyor', [Op Sayısı] = [Op Sayısı] + 1 WHERE [ID] = '" + parcaid + "' ";
                    MessageBox.Show("Parça 'Talaşlı İmalatta' olarak işaretlenmiştir.");
                }
                else if (metroComboBox1.Text.ToString() == "Kaynaklı İmalat")
                {
                    //Talaşlıdan gelen
                    sqldurum = "UPDATE bomlistesi SET [Durum] = '" + bolum + "' , [Tam Op] = [Tam Op] + 1 , [Gerçek Adet] = '" + anlikadet + "' WHERE [ID] = '" + parcaid + "' ";
                    sqlimalatbitti = "UPDATE bomlistesi SET [Kaynak Durum] = 'Devam Ediyor', [Op Sayısı] = [Op Sayısı] + 1 WHERE [ID] = '" + parcaid + "' ";
                    MessageBox.Show("Parça 'Kaynaklı İmalatta' olarak işaretlenmiştir.");
                }
                else if (metroComboBox1.Text.ToString() == "Boya")
                {
                    sqldurum = "UPDATE bomlistesi SET [Durum] = '"+bolum+ "' , [Tam Op] = [Tam Op] + 1 , [Gerçek Adet] = '" + anlikadet + "' WHERE [ID] = '" + parcaid + "' ";
                    sqlimalatbitti = "UPDATE bomlistesi SET [Durum] = 'Boyanacak', [Op Sayısı] = [Op Sayısı] + 1 WHERE [ID] = '" + parcaid + "'  ";
                    MessageBox.Show("Parça 'Boyanacak' olarak işaretlenmiştir.");
                }
                else if (metroComboBox1.Text.ToString() == "Galvaniz")
                {
                    sqldurum = "UPDATE bomlistesi SET [Durum] = '" + bolum + "' , [Tam Op] = [Tam Op] + 1 , [Gerçek Adet] = '" + anlikadet + "' WHERE [ID] = '" + parcaid + "' ";
                    sqlimalatbitti = "UPDATE bomlistesi SET [Durum] = 'Galvanize Gidecek', [Op Sayısı] = [Op Sayısı] + 1 WHERE [ID] = '" + parcaid + "' ";
                    MessageBox.Show("Parça 'Galvanize Gidecek' olarak işaretlenmiştir.");
                }
                else if (metroComboBox1.Text.ToString() == "Kataforez")
                {
                    sqldurum = "UPDATE bomlistesi SET [Durum] = '" + bolum + "' , [Tam Op] = [Tam Op] + 1 , [Gerçek Adet] = '" + anlikadet + "' WHERE [ID] = '" + parcaid + "' ";
                    sqlimalatbitti = "UPDATE bomlistesi SET [Durum] = 'Kataforeze Gidecek', [Op Sayısı] = [Op Sayısı] + 1 WHERE [ID] = '" + parcaid + "' ";
                    MessageBox.Show("Parça 'Kataforeze Gidecek' olarak işaretlenmiştir.");
                }
                else if (metroComboBox1.Text.ToString() == "Fason")
                {
                    sqldurum = "UPDATE bomlistesi SET [Durum] = '" + bolum + "' , [Tam Op] = [Tam Op] + 1 , [Gerçek Adet] = '" + anlikadet + "' WHERE [ID] = '" + parcaid + "' ";
                    sqlimalatbitti = "UPDATE bomlistesi SET [Durum] = 'Tedarikçiye Gidecek', [Op Sayısı] = [Op Sayısı] + 1 WHERE [ID] = '" + parcaid + "' ";
                    MessageBox.Show("Parça 'Tedarikçiye Gidecek' olarak işaretlenmiştir.");
                }
                else if (metroComboBox1.Text.ToString() == "Montaj")
                {
                    sqldurum = "UPDATE bomlistesi SET [Durum] = '" + bolum + "' , [Tam Op] = [Tam Op] + 1 , [Gerçek Adet] = '" + anlikadet + "' WHERE [ID] = '" + parcaid + "' ";
                    sqlimalatbitti = "UPDATE bomlistesi SET [Durum] = 'Montaja Hazır', [Montaj Durum] = 'Devam Ediyor', [Op Sayısı] = [Op Sayısı] + 1 WHERE [ID] = '" + parcaid + "' ";
                    MessageBox.Show("Parça 'Montaja Hazır' olarak işaretlenmiştir.");
                }
                else if (metroComboBox1.Text.ToString() == "Sevkiyat")
                {
                    sqldurum = "UPDATE bomlistesi SET [Durum] = '" + bolum + "' , [Tam Op] = [Tam Op] + 1 , [Gerçek Adet] = '" + anlikadet + "' WHERE [ID] = '" + parcaid + "' ";
                    sqlimalatbitti = "UPDATE bomlistesi SET [Durum] = 'Sevke Hazır' WHERE [ID] = '" + parcaid + "' ";
                    MessageBox.Show("Parça 'Sevke Hazır' olarak işaretlenmiştir.");
                }
                if (parcadurum == "Tedarikçide")
                {
                    string sqlfasonbitti = "UPDATE bomlistesi SET [Fason Durum] = 'Bitti' WHERE [ID] = '" + parcaid + "' ";
                    SQLiteCommand cmdfasonbitti = new SQLiteCommand(sqlfasonbitti, baglan);
                    cmdfasonbitti.ExecuteNonQuery();
                }
                SQLiteCommand cmddurum = new SQLiteCommand(sqldurum, baglan);
                cmddurum.ExecuteNonQuery();
                SQLiteCommand cmdimalatbitti = new SQLiteCommand(sqlimalatbitti, baglan);
                cmdimalatbitti.ExecuteNonQuery();
                SQLiteCommand cmd3 = new SQLiteCommand(sql3, baglan);
                cmd3.ExecuteNonQuery();
                baglan.Close();
                this.Close();
            }
            else if (bolum == "Depoda")
            {
                SQLiteConnection baglan = new SQLiteConnection(baglanti.database);
                baglan.Open();
                sql3 = "INSERT OR IGNORE INTO parcalog([Parça No], [Parça Adedi], [Durum], [Güncelleyen], [Tarih]) VALUES('" + parcanosu + "', '" + parcaadedi + "', '" + bolum + "', '" + usrnme + "', '" + bugun + "') ";
                if (metroComboBox1.Text.ToString() == "")
                {
                    MessageBox.Show("Lütfen bir operasyon seçiniz");
                }
                else if (metroComboBox1.Text.ToString() == "Talaşlı İmalat")
                {
                    sqldurum = "UPDATE bomlistesi SET [Durum] = '" + bolum + "' , [Tam Op] = [Tam Op] + 1 , [Gerçek Adet] = '" + anlikadet + "' WHERE [ID] = '" + parcaid + "' ";
                    sqlimalatbitti = "UPDATE bomlistesi SET [Talaşlı Durum] = 'Devam Ediyor', [Op Sayısı] = [Op Sayısı] + 1 WHERE [ID] = '" + parcaid + "' ";
                    MessageBox.Show("Parça 'Talaşlı İmalatta' olarak işaretlenmiştir.");
                }
                else if (metroComboBox1.Text.ToString() == "Kaynaklı İmalat")
                {
                    sqldurum = "UPDATE bomlistesi SET [Durum] = '" + bolum + "' , [Tam Op] = [Tam Op] + 1 , [Gerçek Adet] = '" + anlikadet + "' WHERE [ID] = '" + parcaid + "' ";
                    sqlimalatbitti = "UPDATE bomlistesi SET [Kaynak Durum] = 'Devam Ediyor', [Op Sayısı] = [Op Sayısı] + 1 WHERE [ID] = '" + parcaid + "' ";
                    MessageBox.Show("Parça 'Kaynaklı İmalatta' olarak işaretlenmiştir.");
                }
                else if (metroComboBox1.Text.ToString() == "Boya")
                {
                    sqldurum = "UPDATE bomlistesi SET [Durum] = '" + bolum + "' , [Tam Op] = [Tam Op] + 1 , [Gerçek Adet] = '" + anlikadet + "' WHERE [ID] = '" + parcaid + "' ";
                    sqlimalatbitti = "UPDATE bomlistesi SET [Durum] = 'Boyanacak', [Op Sayısı] = [Op Sayısı] + 1 WHERE [ID] = '" + parcaid + "'  ";
                    MessageBox.Show("Parça 'Boyanacak' olarak işaretlenmiştir.");
                }
                else if (metroComboBox1.Text.ToString() == "Galvaniz")
                {
                    sqldurum = "UPDATE bomlistesi SET [Durum] = '" + bolum + "' , [Tam Op] = [Tam Op] + 1 , [Gerçek Adet] = '" + anlikadet + "' WHERE [ID] = '" + parcaid + "' ";
                    sqlimalatbitti = "UPDATE bomlistesi SET [Durum] = 'Galvanize Gidecek', [Op Sayısı] = [Op Sayısı] + 1 WHERE [ID] = '" + parcaid + "' ";
                    MessageBox.Show("Parça 'Galvanize Gidecek' olarak işaretlenmiştir.");
                }
                else if (metroComboBox1.Text.ToString() == "Kataforez")
                {
                    sqldurum = "UPDATE bomlistesi SET [Durum] = '" + bolum + "' , [Tam Op] = [Tam Op] + 1 , [Gerçek Adet] = '" + anlikadet + "' WHERE [ID] = '" + parcaid + "' ";
                    sqlimalatbitti = "UPDATE bomlistesi SET [Durum] = 'Kataforeze Gidecek', [Op Sayısı] = [Op Sayısı] + 1 WHERE [ID] = '" + parcaid + "' ";
                    MessageBox.Show("Parça 'Kataforeze Gidecek' olarak işaretlenmiştir.");
                }
                else if (metroComboBox1.Text.ToString() == "Fason")
                {
                    sqldurum = "UPDATE bomlistesi SET [Durum] = '" + bolum + "' , [Tam Op] = [Tam Op] + 1 , [Gerçek Adet] = '" + anlikadet + "' WHERE [ID] = '" + parcaid + "' ";
                    sqlimalatbitti = "UPDATE bomlistesi SET [Durum] = 'Tedarikçiye Gidecek', [Op Sayısı] = [Op Sayısı] + 1 WHERE [ID] = '" + parcaid + "' ";
                    MessageBox.Show("Parça 'Tedarikçiye Gidecek' olarak işaretlenmiştir.");
                }
                else if (metroComboBox1.Text.ToString() == "Montaj")
                {
                    sqldurum = "UPDATE bomlistesi SET [Durum] = '" + bolum + "' , [Tam Op] = [Tam Op] + 1 , [Gerçek Adet] = '" + anlikadet + "' WHERE [ID] = '" + parcaid + "' ";
                    sqlimalatbitti = "UPDATE bomlistesi SET [Durum] = 'Montaja Hazır', [Montaj Durum] = 'Devam Ediyor', [Op Sayısı] = [Op Sayısı] + 1 WHERE [ID] = '" + parcaid + "' ";
                    MessageBox.Show("Parça 'Montaja Hazır' olarak işaretlenmiştir.");
                }
                else if (metroComboBox1.Text.ToString() == "Sevkiyat")
                {
                    sqldurum = "UPDATE bomlistesi SET [Durum] = '" + bolum + "' , [Tam Op] = [Tam Op] + 1 , [Gerçek Adet] = '" + anlikadet + "' WHERE [ID] = '" + parcaid + "' ";
                    sqlimalatbitti = "UPDATE bomlistesi SET [Durum] = 'Sevke Hazır' WHERE [ID] = '" + parcaid + "' ";
                    MessageBox.Show("Parça 'Sevke Hazır' olarak işaretlenmiştir.");
                }
                if (parcadurum == "Tedarikçide")
                {
                    string sqlfasonbitti = "UPDATE bomlistesi SET [Fason Durum] = 'Bitti' WHERE [ID] = '" + parcaid + "' ";
                    SQLiteCommand cmdfasonbitti = new SQLiteCommand(sqlfasonbitti, baglan);
                    cmdfasonbitti.ExecuteNonQuery();
                }
                SQLiteCommand cmddurum = new SQLiteCommand(sqldurum, baglan);
                cmddurum.ExecuteNonQuery();
                SQLiteCommand cmdimalatbitti = new SQLiteCommand(sqlimalatbitti, baglan);
                cmdimalatbitti.ExecuteNonQuery();
                SQLiteCommand cmd3 = new SQLiteCommand(sql3, baglan);
                cmd3.ExecuteNonQuery();
                baglan.Close();
                this.Close();
            }
        }
        private void metroComboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (metroComboBox1.Text.ToString() == "Fason")
            {
                metroLabel3.Visible = true;
                metroTextBox2.Visible = true;
            }
            else
            {
                metroLabel3.Visible = false;
                metroTextBox2.Visible = false;
            }
        }
    }
}