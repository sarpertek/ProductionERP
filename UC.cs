using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using OfficeOpenXml;
using System.Data.SQLite;

namespace _2015
{
    public partial class UC : MetroFramework.Controls.MetroUserControl
    {
        SQLiteConnection baglan { get; set; }
        SQLiteDataAdapter db, dakalite, dadepo;
        DataTable dtuc, dtkalite,dtdepo;
        string durum { get; set; }
        string sqlgonderildi;
        string durumquery;

        private static UC _instance;
        public static UC Instance
        {
            get
            {
                if (_instance == null)
                    _instance = new UC();
                return _instance;
            }
        }
        public UC()
        {
            InitializeComponent();
        }

        private void UC_Load(object sender, EventArgs e)
        {
            metroComboBox1.Items.Add("Depoya Geldi");
            metroComboBox1.Items.Add("Gönderildi");
        }  
        private void metroTextBox3_TextChanged(object sender, EventArgs e)
        {
        }
        private void metroGrid2_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.ColumnIndex == metroGrid2.Columns["Parcano2"].Index)
            {
                try
                {
                    string numara = metroGrid2.Rows[e.RowIndex].Cells[2].Value.ToString();
                    string yil = numara.Substring(0, 2);
                    string sira = numara.Substring(2, 3);
                    string no = (numara + ".slddrw");
                    string dir = @"Z:\PROJE\PROJELER\20%yil%\20%yil%-UC-%sira%\".Replace("%sira%", sira).Replace("%yil%", yil);
                    string[] dirs = Directory.GetFiles(dir, no, SearchOption.AllDirectories);
                    System.Diagnostics.Process.Start(dirs.First());
                }
                catch { MessageBox.Show("Resim Bulunamadı."); }
            }
        }
        private void durumguncellebutton_Click(object sender, EventArgs e)
        {
            int rowindex = metroGrid2.CurrentCell.RowIndex;
            string id = metroGrid2.Rows[rowindex].Cells["ID2"].Value.ToString();
            if (metroComboBox1.Text.ToString() == "Depoya Geldi")
            {
                parcahazir parcahazir = new parcahazir();
                parcahazir.parcaidsi = id;
                parcahazir.bolumidsi = "Depoda";
                parcahazir.parcanosu = metroGrid2.Rows[metroGrid2.CurrentCell.RowIndex].Cells["Parcano2"].Value.ToString();
                parcahazir.parcaadedi = metroGrid2.Rows[metroGrid2.CurrentCell.RowIndex].Cells["toplamadet"].Value.ToString();
                parcahazir.parcadurum = metroGrid2.Rows[metroGrid2.CurrentCell.RowIndex].Cells["Durum2"].Value.ToString();
                parcahazir.ShowDialog();
            }
            else if (metroComboBox1.Text.ToString() == "Gönderildi")
            {
            string anlikdurum = metroGrid2.Rows[rowindex].Cells["Durum2"].Value.ToString();
            SQLiteConnection baglan = new SQLiteConnection(baglanti.database);
            baglan.Open();
                if (anlikdurum == "Galvanize Gidecek")
                {
                    sqlgonderildi = "UPDATE bomlistesi SET [Durum] = 'Galvanizde' , [Tam Op] = [Tam Op] + 1 WHERE [ID] = '" + id + "' ";
                    MessageBox.Show("Parça 'Galvanizde' olarak işaretlenmiştir.");
                    SQLiteCommand cmddurum = new SQLiteCommand(sqlgonderildi, baglan);
                    cmddurum.ExecuteNonQuery();
                }
                if (anlikdurum == "Kataforeze Gidecek")
                {
                    sqlgonderildi = "UPDATE bomlistesi SET [Durum] = 'Kataforezde' , [Tam Op] = [Tam Op] + 1 WHERE [ID] = '" + id + "' ";
                    MessageBox.Show("Parça 'Kataforezde' olarak işaretlenmiştir.");
                    SQLiteCommand cmddurum = new SQLiteCommand(sqlgonderildi, baglan);
                    cmddurum.ExecuteNonQuery();
                }
                if (anlikdurum == "Tedarikçiye Gidecek")
                {
                    sqlgonderildi = "UPDATE bomlistesi SET [Durum] = 'Tedarikçide' , [Tam Op] = [Tam Op] + 1 WHERE [ID] = '" + id + "' ";
                    MessageBox.Show("Parça 'Tedarikçide' olarak işaretlenmiştir.");
                    SQLiteCommand cmddurum = new SQLiteCommand(sqlgonderildi, baglan);
                    cmddurum.ExecuteNonQuery();
                }
            }
            else if(metroComboBox1.Text.ToString() == "")
            {
                MessageBox.Show("Bir durum seçiniz.");
            }
        }
        private void metroButton1_Click(object sender, EventArgs e)
        {
            parcahazir parcahazir = new parcahazir();
            int rowindex = metroGrid1.CurrentCell.RowIndex;
            string id = metroGrid1.Rows[rowindex].Cells["dataGridViewTextBoxColumn5"].Value.ToString();
            parcahazir.parcaidsi = id;
            parcahazir.bolumidsi = "GKK Red";
            parcahazir.parcanosu = metroGrid1.Rows[metroGrid1.CurrentCell.RowIndex].Cells["dataGridViewButtonColumn1"].Value.ToString();
            parcahazir.parcaadedi = metroGrid1.Rows[metroGrid1.CurrentCell.RowIndex].Cells["dataGridViewTextBoxColumn3"].Value.ToString();
            parcahazir.ShowDialog();
            MessageBox.Show("Parça reddedilmiştir.");
            FileInfo newFile = new FileInfo(baglanti.gkkredtemplate);
            ExcelPackage pck = new ExcelPackage(newFile);
            var ws = pck.Workbook.Worksheets.First();
            string parcano = metroGrid1.Rows[metroGrid1.CurrentCell.RowIndex].Cells["dataGridViewButtonColumn1"].Value.ToString();
            string isemri = parcano.Substring(0, 5);
            ws.Cells["E5"].Value = isemri;
            ws.Cells["E7"].Value = metroGrid1.Rows[metroGrid1.CurrentCell.RowIndex].Cells["dataGridViewButtonColumn1"].Value.ToString();
            ws.Cells["E9"].Value = metroGrid1.Rows[metroGrid1.CurrentCell.RowIndex].Cells["dataGridViewTextBoxColumn3"].Value.ToString();
            ws.Cells["N7"].Value = metroGrid1.Rows[metroGrid1.CurrentCell.RowIndex].Cells["dataGridViewTextBoxColumn2"].Value.ToString();
            ws.Cells["E11"].Value = DateTime.Now.ToString("dd/mm/yyyy");
            ws.Cells["E17"].Value = metroGrid1.Rows[metroGrid1.CurrentCell.RowIndex].Cells["dataGridViewTextBoxColumn4"].Value.ToString();
            ws.Cells["AE3"].Value = DateTime.Now.ToString("dd/mm/yyyy");
            baglan = new SQLiteConnection(baglanti.database);
            baglan.Open();
            string sqlbaslik = "SELECT [Proje No], [Firma Adı], [Proje Adı] FROM yuklenmisbom WHERE [Proje No] ='" + isemri + "'";
            using (SQLiteDataAdapter a = new SQLiteDataAdapter(sqlbaslik, baglan))
            {
                DataTable t = new DataTable();
                a.Fill(t);
                if (t.Rows.Count > 0)
                {
                    string firmaadi = t.Rows[0]["Firma Adı"].ToString();
                    string projeadi = t.Rows[0]["Proje Adı"].ToString();
                    //Başlık doldur
                    string baslik = projeadi;
                    ws.Cells["N5"].Value = baslik;
                }
            }
            baglan.Close();
            SaveFileDialog save = new SaveFileDialog();
            save.Filter = "Excel Dosyası|*.xlsx";
            save.OverwritePrompt = true;
            save.CreatePrompt = false;
            if (save.ShowDialog() == DialogResult.OK)
            {
                Stream stream = File.Create(save.FileName);
                pck.SaveAs(stream);
                stream.Close();
                System.Diagnostics.Process.Start(save.FileName);
            }
            urunagacikalite();
        }
        private void metroButton2_Click(object sender, EventArgs e)
        {
            parcahazir parcahazir = new parcahazir();
            int rowindex = metroGrid1.CurrentCell.RowIndex;
            string id = metroGrid1.Rows[rowindex].Cells["dataGridViewTextBoxColumn5"].Value.ToString();
            parcahazir.parcaidsi = id;
            parcahazir.bolumidsi = "GKK Onay";
            parcahazir.parcanosu = metroGrid1.Rows[metroGrid1.CurrentCell.RowIndex].Cells["dataGridViewButtonColumn1"].Value.ToString();
            parcahazir.parcaadedi = metroGrid1.Rows[metroGrid1.CurrentCell.RowIndex].Cells["dataGridViewTextBoxColumn3"].Value.ToString();
            parcahazir.ShowDialog();
            MessageBox.Show("Parça Onaylanmıştır.");
            urunagacikalite();
        }
        private void metroComboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            //metroTextBox3.Text = "";
            //if (metroComboBox2.Text.ToString() == "Tümü")
            //{
            //    sqlquery = "SELECT [ID], [Grup Adı], [Parça No], [Parça Adı], [Toplam Adet], [Kaynaklı İmalat], [Talaşlı İmalat], [Montaj], [Fason], [Lazer], [Durum] FROM bomlistesi";
            //}
            //if (metroComboBox2.Text.ToString() == "Depoda")
            //{
            //    durumquery = "Depoya Geldi";
            //    sqlquery = "SELECT [ID], [Grup Adı], [Parça No], [Parça Adı], [Toplam Adet], [Kaynaklı İmalat], [Talaşlı İmalat], [Montaj], [Fason], [Lazer], [Durum] FROM bomlistesi WHERE [Durum] = '"+ durumquery + "'";
            //}
            //if (metroComboBox2.Text.ToString() == "Kalite Rafında")
            //{
            //    durumquery = "GKK Onay";
            //    sqlquery = "SELECT [ID], [Grup Adı], [Parça No], [Parça Adı], [Toplam Adet], [Kaynaklı İmalat], [Talaşlı İmalat], [Montaj], [Fason], [Lazer], [Durum]  FROM bomlistesi WHERE [Durum] = '" + durumquery + "'";
            //}
            //if (metroComboBox2.Text.ToString() == "İmalatta")
            //{
            //    durumquery = "İmalatta";
            //    sqlquery = "SELECT [ID], [Grup Adı], [Parça No], [Parça Adı], [Toplam Adet], [Kaynaklı İmalat], [Talaşlı İmalat], [Montaj], [Fason], [Lazer], [Durum] FROM bomlistesi WHERE [Durum] = '" + durumquery + "'";
            //}
            //if (metroComboBox2.Text.ToString() == "Kaplamada")
            //{
            //    durumquery = "Kaplamada";
            //    sqlquery = "SELECT [ID], [Grup Adı], [Parça No], [Parça Adı], [Toplam Adet], [Kaynaklı İmalat], [Talaşlı İmalat], [Montaj], [Fason], [Lazer], [Durum] FROM bomlistesi WHERE [Durum] = '" + durumquery + "'";
            //}
            //if (metroComboBox2.Text.ToString() == "Tedarikçide")
            //{
            //    durumquery = "Fasonda";
            //    sqlquery = "SELECT [ID], [Grup Adı], [Parça No], [Parça Adı], [Toplam Adet], [Kaynaklı İmalat], [Talaşlı İmalat], [Montaj], [Fason], [Lazer], [Durum] FROM bomlistesi WHERE [Durum] = '" + durumquery + "'";
            //}
            //baglan = new SQLiteConnection(baglanti.database);
            //baglan.Open();
            //db = new SQLiteDataAdapter(sqlquery,baglan);
            //dtuc = new DataTable();
            //db.Fill(dtuc);
            //metroGrid2.DataSource = dtuc;
            //baglan.Close();
        }
        private void metroTabControl1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (metroTabControl1.SelectedTab == metroTabControl1.TabPages["metrotabpage1"])//your specific tabname
            {
                urunagacikalite();
            }
            if (metroTabControl1.SelectedTab == metroTabControl1.TabPages["metrotabpage1"])//your specific tabname
            {
                urunagacidepo();
            }
        }
        private void urunagacikalite()
        {
            SQLiteConnection baglan = new SQLiteConnection(baglanti.database);
            baglan.Open();
            string sql = "SELECT [ID], [Grup Adı], [Parça No], [Parça Adı], [Toplam Adet], [Kaynak Durum], [Talaşlı Durum], [Montaj Durum], [Fason], [Lazer], [Durum], [Fason Durum] FROM bomlistesi WHERE [Durum] = 'Tedarikçide' OR [Durum] = 'GKK Bekleniyor' ";
            dakalite = new SQLiteDataAdapter(sql, baglan);
            dtkalite = new DataTable();
            dakalite.Fill(dtkalite);
            metroGrid1.DataSource = dtkalite;
            baglan.Close();
        }
        private void urunagacidepo()
        {
            SQLiteConnection baglan = new SQLiteConnection(baglanti.database);
            baglan.Open();
            string sql = "SELECT [ID], [Grup Adı], [Parça No], [Parça Adı], [Toplam Adet], [Kaynak Durum], [Talaşlı Durum], [Montaj Durum], [Fason], [Lazer], [Durum], [Fason Durum] FROM bomlistesi WHERE [Durum] = 'Tedarikçide' OR [Durum] = 'Galvanize Gidecek' OR [Durum] = 'Kataforeze Gidecek' OR [Durum] = 'Tedarikçiye Gidecek' OR [Durum] = 'Kataforezde' OR [Durum] = 'Galvanizde'";
            dadepo = new SQLiteDataAdapter(sql, baglan);
            dtdepo = new DataTable();
            dadepo.Fill(dtdepo);
            metroGrid2.DataSource = dtdepo;
            baglan.Close();
        }
    }
    }

