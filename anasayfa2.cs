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
    public partial class anasayfa2 : MetroFramework.Controls.MetroUserControl
    {
        string StringA { get; set; }
        string StringB { get; set; }
        string projeno { get; set; }
        int queryrow { get; set; }
        int ilerlenen, ilerlenecek,saat;
        string sqlmontajdurum;
        DataTable dtproje, dtkaynak, dttalasli, dtmontaj, dtfason,dtozet,dtfasongecmisi;
        SQLiteDataAdapter da, dakaynak, datalasli, damontaj, dafason,daozet,dafasongecmisi;

        private static anasayfa2 _instance;
        public static anasayfa2 Instance
        {
            get
            {
                if (_instance == null)
                    _instance = new anasayfa2();
                return _instance;
            }
        }
        public anasayfa2()
        {
            InitializeComponent();
            metroProgressBar1.Visible = false;
        }

        private void anasayfa2_Load(object sender, EventArgs e)
        {     
            tumbomlar();
        }
        private void tumbomlar()
        {
            SQLiteConnection baglan = new SQLiteConnection(baglanti.database);
            baglan.Open();
            string sql = "SELECT *FROM yuklenmisbom";
            SQLiteDataAdapter da = new SQLiteDataAdapter(sql, baglan);
            DataTable dtbomlar = new DataTable();
            da.Fill(dtbomlar);
            metroGridbomlar.DataSource = dtbomlar;
            baglan.Close();
        }
        private void metroTile2_Click(object sender, EventArgs e)
        {
            try
            {
                string usrnme = Environment.UserName;
                SQLiteConnection baglan = new SQLiteConnection(baglanti.database);
                OpenFileDialog openfiledialog1 = new OpenFileDialog();
                openfiledialog1.Filter = "Excel Dosyaları|*.xls;*.xlsx;*.xlsm";
                if (openfiledialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    string bomlistesi = openfiledialog1.FileName;
                    string bomlistesiadi = openfiledialog1.SafeFileName;
                    baglan.Open();
                    FileInfo newFile = new FileInfo(bomlistesi);
                    ExcelPackage pck = new ExcelPackage(newFile);
                    var ws = pck.Workbook.Worksheets.First();

                    string firmaadi = ws.Cells[1, 8].Value.ToString();
                    string projeadi = ws.Cells[2, 8].Value.ToString();
                    string grupadi = ws.Cells[5, 8].Value.ToString();
                    string[] isemric = bomlistesiadi.Split('_');
                    string isemrinosu = isemric.First();
                    string sql = "SELECT *FROM yuklenmisbom WHERE [Firma Adı]=@firmaadi AND [Proje Adı]=@projeadi AND [Grup Adı]=@grupadi AND [Proje No]=@projeno";
                    SQLiteParameter prm2 = new SQLiteParameter("projeadi", projeadi);
                    SQLiteParameter prm1 = new SQLiteParameter("firmaadi", firmaadi);
                    SQLiteParameter prm3 = new SQLiteParameter("grupadi", grupadi);
                    SQLiteParameter prm4 = new SQLiteParameter("projeno", isemrinosu);
                    SQLiteCommand cmd = new SQLiteCommand(sql, baglan);
                    cmd.Parameters.Add(prm1);
                    cmd.Parameters.Add(prm2);
                    cmd.Parameters.Add(prm3);
                    cmd.Parameters.Add(prm4);
                    DataTable dt1 = new DataTable();
                    SQLiteDataAdapter da = new SQLiteDataAdapter(cmd);
                    da.Fill(dt1);

                    if (dt1.Rows.Count == 0)
                    {
                        //Tümünü excelden al                  
                        metroProgressBar1.Visible = true;
                        metroProgressBar1.Value = 0;
                        DataTable tbl = new DataTable();
                        // header ara
                        var query1 = (from cell in ws.Cells[2,2,11,11]
                                      where cell.Value != null && cell.Value.ToString() == "PARCA NO"
                                      select cell);
                        foreach (var cell in query1)
                        {
                            queryrow = cell.Start.Row;
                        }


                    bool hasHeader = true;
                    int sRow = queryrow + 1;
                    foreach (var firstRowCell in ws.Cells[queryrow, 1, queryrow, 12])
                    {
                        tbl.Columns.Add(hasHeader ? firstRowCell.Text : string.Format("Column {0}", firstRowCell.Start.Column));
                    }
                    var startRow = hasHeader ? sRow : 1;
                    for (var rowNum = startRow; rowNum <= ws.Dimension.End.Row; rowNum++)
                    {
                        if (ws.Cells[rowNum, 2].Value != null && ws.Cells[rowNum, 2].Value.ToString().Length == 12)
                        {
                            var wsRow = ws.Cells[rowNum, 1, rowNum, 12];
                            var row = tbl.NewRow();
                            foreach (var cell in wsRow)
                            {
                                row[cell.Start.Column - 1] = cell.Text;
                            }
                            tbl.Rows.Add(row);
                        }
                        else
                        { }
                    }
                    DataView datafilter = tbl.DefaultView;
                        datafilter.RowFilter = "ACIKLAMA NOT LIKE 'HAZIR*'";
                        DataTable aratablo = datafilter.ToTable();
                        string bugun = DateTime.Now.ToString("dd/MM/yyyy");
                        string sql1 = "INSERT INTO yuklenmisbom([Firma Adı], [Proje Adı], [Proje No], [Eklenme Tarihi], [Grup Adı], [Ekleyen]) Values('" + firmaadi + "', '" + projeadi + "', '" + isemrinosu + "','" + bugun + "', '" + grupadi + "', '" + usrnme + "') ";
                        SQLiteCommand cmd1 = new SQLiteCommand(sql1, baglan);
                        cmd1.ExecuteNonQuery();
                        metroProgressBar1.Maximum = aratablo.Rows.Count;
                        for (int i = 0; i < aratablo.Rows.Count; i++)
                        {
                            metroProgressBar1.Value += 1;
                            string denemedurumu = "Bilgi Bekleniyor";
                            string zero = "0";
                            string sql2 = "INSERT OR IGNORE INTO bomlistesi([Proje Numarası], [Parça No], [Parça Adı], [Grup Adı], [Durum], [Toplam Adet], [Birim Fiyat], [Toplam], [Op Sayısı], [Tam Op], [Kaynaklı İmalat], [Talaşlı İmalat], [Montaj], [Montaj2], [Lazer]) VALUES('" + isemrinosu.ToString() + "', '" + aratablo.Rows[i]["PARCA NO"].ToString().Trim() + "' , '" + aratablo.Rows[i]["PARCA ADI"].ToString().Trim() + "' , '" + grupadi + "' , '" + denemedurumu + "', '" + aratablo.Rows[i][4].ToString().Trim() + "', '"+zero+"', '"+zero+ "','1', '0', 'YOK', 'YOK', 'YOK', 'YOK' , 'YOK')";
                            SQLiteCommand cmd2 = new SQLiteCommand(sql2, baglan);
                            cmd2.ExecuteNonQuery();
                            string sql3 = "INSERT OR IGNORE INTO parcalog([Parça No], [Parça Adedi], [Durum], [Güncelleyen], [Tarih]) VALUES('" + aratablo.Rows[i]["PARCA NO"].ToString().Trim() + "', '" + aratablo.Rows[i][4].ToString().Trim() + "', '" + denemedurumu + "', '" + usrnme + "', '" + bugun + "') ";
                            SQLiteCommand cmd3 = new SQLiteCommand(sql3, baglan);
                            cmd3.ExecuteNonQuery();
                        }
                        baglan.Close();
                        metroProgressBar1.Visible = false;
                        MessageBox.Show("Bom listesi başarıyla eklendi.");
                    }
                    else
                    {
                        MessageBox.Show("Bu bom listesi daha önce eklenmiş.");
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Bir hata oluştu lütfen tekrar deneyiniz.");
                MessageBox.Show(ex.Message);
            }
            tumbomlar();
        }


        private void anasayfa2_FormClosing(object sender, FormClosingEventArgs e)
        {
            Application.Exit();
        }

        private void metroTile3_Click(object sender, EventArgs e)
        {
            tumbomlar();
        }

        private void metroTextBox1_TextChanged(object sender, EventArgs e)
        {
            if (metroTextBox1.Text.ToString() != "")
            {
                string ara = metroTextBox1.Text.ToString();
                SQLiteConnection baglan = new SQLiteConnection(baglanti.database);
                baglan.Open();

                string sql = "SELECT *FROM yuklenmisbom WHERE [Proje No] LIKE '%" + ara + "%' OR [Firma Adı] LIKE '%" + ara + "%' OR [Grup Adı] LIKE '%" + ara + "%' OR [Ekleyen] LIKE '%" + ara + "%' OR [Proje Adı] LIKE '%" + ara + "%' ";
                SQLiteDataAdapter da = new SQLiteDataAdapter(sql, baglan);
                DataTable dtbomlarfiltered = new DataTable();
                da.Fill(dtbomlarfiltered);
                metroGridbomlar.DataSource = dtbomlarfiltered;
                baglan.Close();
            }
            else { tumbomlar(); }
        }

        private void urunagaci()
        {
            projeno = metroTextBox3.Text;
            SQLiteConnection baglan = new SQLiteConnection(baglanti.database);
            baglan.Open();
            string sql = "SELECT [ID], [Grup Adı], [Parça No], [Parça Adı], [Toplam Adet], [Kaynaklı İmalat], [Talaşlı İmalat], [Montaj], [Montaj2], [Fason], [Lazer], [Ust Grup], [Birim Fiyat], [Toplam], [Durum] FROM bomlistesi WHERE [Proje Numarası]='" + projeno + "'";
            da = new SQLiteDataAdapter(sql, baglan);
            dtproje = new DataTable();
            da.Fill(dtproje);
            metroGrid2.DataSource = dtproje;

            //baglan.Close();

            ////baslik doldur
            //baglan.Open();
            string sqlbaslik = "SELECT [Proje No], [Firma Adı], [Proje Adı] FROM yuklenmisbom WHERE [Proje No] ='" + projeno + "'";
            using (SQLiteDataAdapter a = new SQLiteDataAdapter(sqlbaslik, baglan))
            {
                DataTable t = new DataTable();
                a.Fill(t);
                if (t.Rows.Count > 0)
                {
                    string firmaadi = t.Rows[0]["Firma Adı"].ToString();
                    string projeadi = t.Rows[0]["Proje Adı"].ToString();
                    //Başlık doldur
                    string baslik = projeno + " - " + firmaadi + " - " + projeadi;
                    metroLabel1.Text = baslik;
                }
            }
            baglan.Close();
        }
        private void urunagacikaynakli()
        {
            SQLiteConnection baglan = new SQLiteConnection(baglanti.database);
            baglan.Open();
            string sql = "SELECT [ID], [Grup Adı], [Parça No], [Parça Adı], [Toplam Adet], [Kaynak Durum], [Talaşlı Durum], [Montaj Durum], [Fason], [Lazer], [Durum] FROM bomlistesi WHERE [Kaynak Durum] = 'Devam Ediyor' OR [Kaynak Durum] = 'Fason Talep Ediyor' ";
            dakaynak = new SQLiteDataAdapter(sql, baglan);
            dtkaynak = new DataTable();
            dakaynak.Fill(dtkaynak);
            metroGrid1.DataSource = dtkaynak;
            baglan.Close();
        }
        private void urunagacitalasli()
        {
            SQLiteConnection baglan = new SQLiteConnection(baglanti.database);
            baglan.Open();
            string sql = "SELECT [ID], [Grup Adı], [Parça No], [Parça Adı], [Toplam Adet], [Kaynak Durum], [Talaşlı Durum], [Montaj Durum], [Fason], [Lazer], [Durum] FROM bomlistesi WHERE [Talaşlı Durum] = 'Devam Ediyor' OR [Talaşlı Durum] = 'Fason Talep Ediyor' ";
            datalasli = new SQLiteDataAdapter(sql, baglan);
            dttalasli = new DataTable();
            datalasli.Fill(dttalasli);
            metroGrid3.DataSource = dttalasli;
            baglan.Close();
        }
        private void urunagacimontaj()
        {
            SQLiteConnection baglan = new SQLiteConnection(baglanti.database);
            baglan.Open();
            string sql = "SELECT [ID], [Grup Adı], [Parça No], [Parça Adı], [Toplam Adet], [Kaynak Durum], [Talaşlı Durum], [Montaj Durum], [Fason], [Lazer], [Durum] FROM bomlistesi WHERE [Montaj Durum] = 'Devam Ediyor' OR [Montaj Durum] = 'Montaj Bekleniyor' ";
            damontaj = new SQLiteDataAdapter(sql, baglan);
            dtmontaj = new DataTable();
            damontaj.Fill(dtmontaj);
            metroGrid5.DataSource = dtmontaj;
            baglan.Close();
        }
        private void metroButton1_Click(object sender, EventArgs e)
        {
            urunagaci();
            try
            {
                metroComboBox1.Items.Clear();
                metroComboBox1.Items.Add("Tümü");
                for (int intCount = 0; intCount < dtproje.Rows.Count; intCount++)
                {
                    if (!metroComboBox1.Items.Contains(dtproje.Rows[intCount]["Grup Adı"].ToString()))
                    {
                        metroComboBox1.Items.Add(dtproje.Rows[intCount]["Grup Adı"].ToString());
                    }
                }
                metroComboBox1.Text = "Tümü";
            }
            catch { }
            
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

        private void metroTextBox4_TextChanged(object sender, EventArgs e)
        {
            if (metroTextBox4.Text.ToString() != "")
            {
                string ara = metroTextBox4.Text.ToString();
                SQLiteConnection baglan = new SQLiteConnection(baglanti.database);
                baglan.Open();
                string sql = "SELECT [ID], [Grup Adı], [Parça No], [Parça Adı], [Toplam Adet], [Kaynaklı İmalat], [Talaşlı İmalat], [Montaj], [Fason], [Lazer], [Birim Fiyat], [Toplam], [Durum] FROM bomlistesi WHERE [Proje Numarası]='" + projeno + "' AND ([Grup Adı] LIKE '%" + ara + "%' OR [Parça No] LIKE '%" + ara + "%' OR [Parça Adı] LIKE '%" + ara + "%')";
                da = new SQLiteDataAdapter(sql, baglan);
                dtproje = new DataTable();
                da.Fill(dtproje);
                metroGrid2.DataSource = dtproje;
                baglan.Close();
            }
            else if (metroComboBox1.Text == "Tümü")
            { urunagaci(); }
            else {
                string ara = metroComboBox1.Text.ToString();
                SQLiteConnection baglan = new SQLiteConnection(baglanti.database);
                baglan.Open();
                string sql = "SELECT [ID], [Grup Adı], [Parça No], [Parça Adı], [Toplam Adet], [Kaynaklı İmalat], [Talaşlı İmalat], [Montaj], [Fason], [Lazer], [Birim Fiyat], [Toplam], [Durum] FROM bomlistesi WHERE [Grup Adı] LIKE '%" + ara + "%' ";
                da = new SQLiteDataAdapter(sql, baglan);
                dtproje = new DataTable();
                da.Fill(dtproje);
                metroGrid2.DataSource = dtproje;
                baglan.Close();
            }
        }

        private void metroTile1_Click(object sender, EventArgs e)
        {
            SQLiteConnection baglan = new SQLiteConnection(baglanti.database);
            baglan.Open();
            int rowindex = metroGridbomlar.CurrentCell.RowIndex;
            string id = metroGridbomlar.Rows[rowindex].Cells[idbom.Index].Value.ToString();
            string projenosu = metroGridbomlar.Rows[rowindex].Cells["Proje"].Value.ToString();
            string grupadisi = metroGridbomlar.Rows[rowindex].Cells["GrupAdi"].Value.ToString();
            string sql = "DELETE FROM yuklenmisbom WHERE ID = '" + id + "' ";
            SQLiteCommand cmd = new SQLiteCommand(sql, baglan);
            cmd.ExecuteNonQuery();
            string sql2 = "DELETE FROM bomlistesi WHERE [Proje Numarası] = '" + projenosu + "' AND [Grup Adı] = '" + grupadisi + "' ";
            SQLiteCommand cmd2 = new SQLiteCommand(sql2, baglan);
            cmd2.ExecuteNonQuery();
            baglan.Close();
            tumbomlar();
        }

        private void metroButton1_Click_1(object sender, EventArgs e)
        {
            DataTable dtaktar = new DataTable();
            foreach (DataGridViewColumn col in metroGrid2.Columns)
            {
                dtaktar.Columns.Add(col.HeaderText);
            }

            foreach (DataGridViewRow row in metroGrid2.Rows)
            {
                DataRow dRow = dtaktar.NewRow();
                foreach (DataGridViewCell cell in row.Cells)
                {
                    dRow[cell.ColumnIndex] = cell.Value;
                }
                dtaktar.Rows.Add(dRow);
            }

            FileInfo newFile = new FileInfo(baglanti.fasonlistesi);
            ExcelPackage pck = new ExcelPackage(newFile);
            //Add the Content sheet
            var ws = pck.Workbook.Worksheets.First();
            ws.View.ShowGridLines = true;
            //Başlık
            ws.Cells["A7"].LoadFromDataTable(dtaktar, true);
            ws.Cells["A7:H7"].Style.Font.Bold = true;
            ws.Cells["A7:H7"].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            ws.Cells["A7:H7"].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Black);
            ws.Cells["A7:H7"].Style.Font.Color.SetColor(Color.White);
            int sonrow = ws.Dimension.End.Row + 1;
            ws.Cells["h:h"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
            ws.Cells["A7:H7"].AutoFilter = true;
            ws.Cells.AutoFitColumns(0);  //Autofit columns for all cells
            ws.DeleteColumn(1);
            string path = @"C:\Users\uretim6\Desktop\" + "deneme fason.xlsx";
            Stream stream = File.Create(path);
            pck.SaveAs(stream);
            stream.Close();
        }

        private void metroButton2_Click(object sender, EventArgs e)
        {
            DialogResult dialogResult = MessageBox.Show("Parçanın TAMAMINI mı fasona vermek istiyorsunuz?", "Uyarı", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (dialogResult == DialogResult.Yes)
            {
                int rowindex = metroGrid1.CurrentCell.RowIndex;
                if (metroGrid1.Rows[rowindex].Cells["dataGridViewTextBoxColumn8"].Value.ToString() != "Onay Bekliyor")
                {
                    SQLiteConnection baglan = new SQLiteConnection(baglanti.database);
                    baglan.Open();
                    //KAYNAKLI EKLE    
                    string id = metroGrid1.Rows[rowindex].Cells[dataGridViewTextBoxColumn5.Index].Value.ToString();
                    string sqldurum = "UPDATE bomlistesi SET [Fason] = 'Fason', [Kaynak Durum] = 'Fason Talep Ediyor' , [Fason Miktarı] = 'Tamamı' , [Durum] = 'Onay Bekliyor' WHERE [ID] = '" + id + "' ";
                    SQLiteCommand cmddurum = new SQLiteCommand(sqldurum, baglan);
                    cmddurum.ExecuteNonQuery();
                    baglan.Close();
                    urunagacikaynakli();
                    MessageBox.Show("Fason talebiniz işleme alınmıştır.");
                }
                else { MessageBox.Show("Talebiniz daha önce işleme alınmış..."); }
            }
            else if (dialogResult == DialogResult.No)
            {
                int rowindex = metroGrid1.CurrentCell.RowIndex;
                if (metroGrid1.Rows[rowindex].Cells["dataGridViewTextBoxColumn8"].Value.ToString() != "Onay Bekliyor")
                {
                    SQLiteConnection baglan = new SQLiteConnection(baglanti.database);
                    baglan.Open();
                    //KAYNAKLI EKLE    
                    string id = metroGrid1.Rows[rowindex].Cells[dataGridViewTextBoxColumn5.Index].Value.ToString();
                    string sqldurum = "UPDATE bomlistesi SET [Fason] = 'Fason', [Kaynak Durum] = 'Fason Talep Ediyor' , [Fason Miktarı] = 'Kısmi' , [Durum] = 'Onay Bekliyor' WHERE [ID] = '" + id + "' ";
                    SQLiteCommand cmddurum = new SQLiteCommand(sqldurum, baglan);
                    cmddurum.ExecuteNonQuery();
                    baglan.Close();
                    urunagacikaynakli();
                    MessageBox.Show("Fason talebiniz işleme alınmıştır.");
                }
                else { MessageBox.Show("Talebiniz daha önce işleme alınmış..."); }
            }
            
        }

        private void metroButton5_Click(object sender, EventArgs e)
        {
            Talep_Onayı talep_onayı = new Talep_Onayı();
            int rowindex = metroGrid4.CurrentCell.RowIndex;
            string a = metroGrid4.Rows[rowindex].Cells["dataGridViewButtonColumn3"].Value.ToString();
            string b = metroGrid4.Rows[rowindex].Cells["dataGridViewTextBoxColumn17"].Value.ToString();
            string c = metroGrid4.Rows[rowindex].Cells["dataGridViewTextBoxColumn18"].Value.ToString();
            string d = metroGrid4.Rows[rowindex].Cells["dataGridViewTextBoxColumn21"].Value.ToString();
            string toplamadet = metroGrid4.Rows[rowindex].Cells["dataGridViewTextBoxColumn19"].Value.ToString();
            talep_onayı.fasonparcastring = a;
            talep_onayı.fasonparcastring2 = b;
            talep_onayı.fasonparcastring3 = c;
            talep_onayı.fasonparcaid = d;
            talep_onayı.fasonparcaadedistring = toplamadet;
            talep_onayı.ShowDialog();
            fasontalepleri();
        }

        private void metroGrid4_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            //int rowindex = e.RowIndex;
            //SQLiteConnection baglan = new SQLiteConnection(baglanti.database);
            //baglan.Open();
            ////KAYNAKLI EKLE   
            //int rowindex = metroGrid1.CurrentCell.RowIndex;
            //string id = metroGrid1.Rows[rowindex].Cells[dataGridViewTextBoxColumn5.Index].Value.ToString();
            //MessageBox.Show(id);
            //string sqldurum = "UPDATE bomlistesi SET [Fason] = 'Fason' , [Durum] = 'Onay Bekliyor' WHERE [ID] = '" + id + "' ";
            //SQLiteCommand cmddurum = new SQLiteCommand(sqldurum, baglan);
            //cmddurum.ExecuteNonQuery();
            //baglan.Close();
            //urunagacikaynakli();
        }

        private void metroButton8_Click(object sender, EventArgs e)
        {
            parcahazir parcahazir = new parcahazir();
            int rowindex = metroGrid1.CurrentCell.RowIndex;
            string id = metroGrid1.Rows[rowindex].Cells[dataGridViewTextBoxColumn5.Index].Value.ToString();
            parcahazir.parcaidsi = id;
            parcahazir.bolumidsi = "Kaynak Durum";
            parcahazir.ShowDialog();
        }

        private void metroButton9_Click(object sender, EventArgs e)
        {
            parcahazir parcahazir = new parcahazir();
            int rowindex = metroGrid3.CurrentCell.RowIndex;
            string id = metroGrid3.Rows[rowindex].Cells["dataGridViewTextBoxColumn13"].Value.ToString();
            parcahazir.parcaidsi = id;
            parcahazir.bolumidsi = "Talaşlı Durum";
            parcahazir.ShowDialog();
        }

        private void metroButton4_Click(object sender, EventArgs e)
        {
            DialogResult dialogResult = MessageBox.Show("Parçanın TAMAMINI mı fasona vermek istiyorsunuz?", "Uyarı", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (dialogResult == DialogResult.Yes)
            {
                int rowindex = metroGrid3.CurrentCell.RowIndex;
                if (metroGrid3.Rows[rowindex].Cells["dataGridViewTextBoxColumn16"].Value.ToString() != "Onay Bekliyor")
                {
                    SQLiteConnection baglan = new SQLiteConnection(baglanti.database);
                    baglan.Open();
                    //TALAŞLI EKLE   
                    string id = metroGrid3.Rows[rowindex].Cells[dataGridViewTextBoxColumn13.Index].Value.ToString();
                    string sqldurum = "UPDATE bomlistesi SET [Fason] = 'Fason', [Talaşlı Durum] = 'Fason Talep Ediyor', [Fason Miktarı] = 'Tamamı' , [Durum] = 'Onay Bekliyor' WHERE [ID] = '" + id + "' ";
                    SQLiteCommand cmddurum = new SQLiteCommand(sqldurum, baglan);
                    cmddurum.ExecuteNonQuery();
                    baglan.Close();
                    urunagacitalasli();
                    MessageBox.Show("Fason talebiniz işleme alınmıştır.");
                }
                else { MessageBox.Show("Talebiniz daha önce işleme alınmış..."); }
            }
            else if (dialogResult == DialogResult.No)
            {
                int rowindex = metroGrid3.CurrentCell.RowIndex;
                if (metroGrid3.Rows[rowindex].Cells["dataGridViewTextBoxColumn16"].Value.ToString() != "Onay Bekliyor")
                {
                    SQLiteConnection baglan = new SQLiteConnection(baglanti.database);
                    baglan.Open();
                    //TALAŞLI EKLE    
                    string id = metroGrid3.Rows[rowindex].Cells[dataGridViewTextBoxColumn13.Index].Value.ToString();
                    string sqldurum = "UPDATE bomlistesi SET [Fason] = 'Fason', [Talaşlı Durum] = 'Fason Talep Ediyor', [Fason Miktarı] = 'Kısmi', [Durum] = 'Onay Bekliyor' WHERE [ID] = '" + id + "' ";
                    SQLiteCommand cmddurum = new SQLiteCommand(sqldurum, baglan);
                    cmddurum.ExecuteNonQuery();
                    baglan.Close();
                    urunagacitalasli();
                    MessageBox.Show("Fason talebiniz işleme alınmıştır.");
                }
                else { MessageBox.Show("Talebiniz daha önce işleme alınmış..."); }
            }
        }

        private void metroButton10_Click(object sender, EventArgs e)
        {
            int rowindex = metroGrid4.CurrentCell.RowIndex;
            //KAYNAK
            if (metroGrid4.Rows[rowindex].Cells["Column2"].Value.ToString() == "Fason Talep Ediyor")
            {
                string id = metroGrid4.Rows[rowindex].Cells[dataGridViewTextBoxColumn21.Index].Value.ToString();
                SQLiteConnection baglan = new SQLiteConnection(baglanti.database);
                baglan.Open();
                string sqlfasonbom = "UPDATE bomlistesi SET [Fason] = '' , [Durum] = 'İmalatta', [Kaynak Durum] = 'Devam Ediyor', [Fason Durum] = 'Bitti' WHERE [ID] = '" + id + "' ";
                SQLiteCommand cmdfasonbom = new SQLiteCommand(sqlfasonbom, baglan);
                cmdfasonbom.ExecuteNonQuery();
                baglan.Close();
                MessageBox.Show("Fason talebi reddedilmiştir.");
            }
            //TALAŞLI
            if (metroGrid4.Rows[rowindex].Cells["Column3"].Value.ToString() == "Fason Talep Ediyor")
            {
                string id = metroGrid4.Rows[rowindex].Cells[dataGridViewTextBoxColumn21.Index].Value.ToString();
                SQLiteConnection baglan = new SQLiteConnection(baglanti.database);
                baglan.Open();
                string sqlfasonbom = "UPDATE bomlistesi SET [Fason] = '' , [Durum] = 'İmalatta', [Talaşlı Durum] = 'Devam Ediyor', [Fason Durum] = 'Bitti' WHERE [ID] = '" + id + "' ";
                SQLiteCommand cmdfasonbom = new SQLiteCommand(sqlfasonbom, baglan);
                cmdfasonbom.ExecuteNonQuery();
                baglan.Close();
                MessageBox.Show("Fason talebi reddedilmiştir.");
            }
        }

        private void metroGrid1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.ColumnIndex == metroGrid1.Columns["dataGridViewButtonColumn1"].Index)
            {
                try
                {
                    string numara = metroGrid1.Rows[e.RowIndex].Cells[2].Value.ToString();
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

        private void metroButton16_Click(object sender, EventArgs e)
        {
            projeozetiliste();
            projeozetirenk();
            try
            {
                metroComboBox5.Items.Clear();
                metroComboBox5.Items.Add("Tümü");
                for (int intCount = 0; intCount < dtozet.Rows.Count; intCount++)
                {
                    if (!metroComboBox5.Items.Contains(dtozet.Rows[intCount]["Grup Adı"].ToString()))
                    {
                        metroComboBox5.Items.Add(dtozet.Rows[intCount]["Grup Adı"].ToString());
                    }
                }
                metroComboBox5.Text = "Tümü";
            }
            catch { }
        }

        private void metroGrid6_SelectionChanged(object sender, EventArgs e)
        {
            metroGrid6.ClearSelection();
        }

        private void metroTextBox2_TextChanged(object sender, EventArgs e)
        {
            if (metroTextBox2.Text.ToString() != "")
            {
                string ara = metroTextBox2.Text.ToString();
                SQLiteConnection baglan = new SQLiteConnection(baglanti.database);
                baglan.Open();
                string sql = "SELECT [ID], [Grup Adı], [Parça No], [Parça Adı], [Toplam Adet], [Kaynak Durum], [Talaşlı Durum], [Montaj Durum], [Fason], [Lazer], [Durum] FROM bomlistesi WHERE ([Kaynak Durum] = 'Devam Ediyor' OR [Kaynak Durum] = 'Fason Talep Ediyor') AND ([Proje Numarası] LIKE '%" + ara + "%' OR [Parça No] LIKE '%" + ara + "%' OR [Parça Adı] LIKE '%" + ara + "%')";
                dakaynak = new SQLiteDataAdapter(sql, baglan);
                dtkaynak = new DataTable();
                dakaynak.Fill(dtkaynak);
                metroGrid1.DataSource = dtkaynak;
                baglan.Close();
            }
            else
            {
                urunagacikaynakli();
            }
        }

        private void metroTextBox10_TextChanged(object sender, EventArgs e)
        {
            if (metroTextBox10.Text.ToString() != "")
            {
                string ara = metroTextBox10.Text.ToString();
                SQLiteConnection baglan = new SQLiteConnection(baglanti.database);
                baglan.Open();
                string sql = "SELECT [fasonid], [Parça No], [Grup Adı],[Parça Adı], [Sipariş Tarihi], [Termin], [Sipariş Adedi], [Birim Fiyat], [Toplam], [Onaylayan], [Fason], [İşçilik], [Malzeme], [Kaplama] FROM fasontakibi WHERE [Parça No] LIKE '%" + ara + "%' OR [Parça Adı] LIKE '%" + ara + "%' OR [Fason] LIKE '%" + ara + "%'";
                dafasongecmisi = new SQLiteDataAdapter(sql, baglan);
                dtfasongecmisi = new DataTable();
                dafasongecmisi.Fill(dtfasongecmisi);
                metroGrid7.DataSource = dtfasongecmisi;
                baglan.Close();
            }
            else
            {
                fasonimalatgecmisi();
            }
        }

        private void metroGrid7_SelectionChanged(object sender, EventArgs e)
        {
            metroGrid7.ClearSelection();
        }

        private void metroGrid7_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.ColumnIndex == metroGrid7.Columns["dataGridViewButtonColumn7"].Index)
            {
                try
                {
                    string numara = metroGrid7.Rows[e.RowIndex].Cells["dataGridViewButtonColumn7"].Value.ToString();
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

        private void metroComboBox5_SelectedIndexChanged(object sender, EventArgs e)
        {
            metroTextBox8.Text = "";
            if (metroComboBox5.Text.ToString() == "Tümü")
            {
                projeozetiliste();
                projeozetirenk();
                
            }
            else
            {        
                string ara = metroComboBox5.Text.ToString();
                SQLiteConnection baglan = new SQLiteConnection(baglanti.database);
                baglan.Open();
                string sql = "SELECT [ID], [Grup Adı], [Parça No], [Parça Adı], [Toplam Adet], [Kaynak Durum], [Talaşlı Durum], [Montaj Durum], [Fason], [Fason Durum], [Lazer], [Birim Fiyat], [Toplam], [Durum] FROM bomlistesi WHERE [Grup Adı] LIKE '%" + ara + "%'";
                daozet = new SQLiteDataAdapter(sql, baglan);
                dtozet = new DataTable();
                daozet.Fill(dtozet);
                metroGrid6.DataSource = dtozet;
                baglan.Close();
                projeozetirenk();
            }
        }

        private void metroTextBox8_TextChanged(object sender, EventArgs e)
        {
            if (metroTextBox8.Text.ToString() != "")
            {
                projeno = metroTextBox9.Text;
                string ara = metroTextBox8.Text.ToString();
                SQLiteConnection baglan = new SQLiteConnection(baglanti.database);
                baglan.Open();
                string sqlozetara = "SELECT [ID], [Grup Adı], [Parça No], [Parça Adı], [Toplam Adet], [Kaynak Durum], [Talaşlı Durum], [Montaj Durum], [Fason], [Fason Durum], [Lazer], [Birim Fiyat], [Toplam], [Durum] FROM bomlistesi WHERE [Proje Numarası]='" + projeno + "' AND ([Parça No] LIKE '%" + ara + "%' OR [Parça Adı] LIKE '%" + ara + "%')";
                daozet = new SQLiteDataAdapter(sqlozetara, baglan);
                dtozet = new DataTable();
                daozet.Fill(dtozet);
                metroGrid6.DataSource = dtozet;
                baglan.Close();
                projeozetirenk();
            }
            else
            {
                projeozetiliste();
                projeozetirenk();
            }
        }
        private void metroTextBox9_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                metroButton16.PerformClick();
            }
        }

        private void metroGrid6_Sorted(object sender, EventArgs e)
        {
            projeozetirenk();
        }

        private void metroGrid5_Sorted(object sender, EventArgs e)
        {
            montajrenk();
        }

        private void metroGrid5_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.ColumnIndex == metroGrid5.Columns["dataGridViewButtonColumn4"].Index)
            {
                try
                {
                    string numara = metroGrid5.Rows[e.RowIndex].Cells["dataGridViewButtonColumn4"].Value.ToString();
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

        private void metroButton11_Click(object sender, EventArgs e)
        {
            int row = metroGrid5.CurrentCell.RowIndex;
            if (metroGrid5.Rows[row].Cells["dataGridViewTextBoxColumn26"].Value.ToString() == "Montaja Hazır")
            {
                SQLiteConnection baglan = new SQLiteConnection(baglanti.database);
                baglan.Open();
                string parcaid = metroGrid5.Rows[row].Cells["dataGridViewTextBoxColumn25"].Value.ToString();
                string a = metroGrid5.Rows[row].Cells["dataGridViewButtonColumn4"].Value.ToString().Substring(metroGrid5.Rows[row].Cells["dataGridViewButtonColumn4"].Value.ToString().Length - 2);
                string parcanoustgrup = metroGrid5.Rows[row].Cells["dataGridViewButtonColumn4"].Value.ToString();
                string ustgrup = parcanoustgrup.Remove(parcanoustgrup.Length - 2, 2) + "00";
                if (a == "00")
                {
                    sqlmontajdurum = "UPDATE bomlistesi SET [Durum] = 'Sevke Hazır', [Montaj Durum] = 'Bitti' , [Tam Op] = [Tam Op] + 1 WHERE [ID] = '" + parcaid + "' ";
                    string sqlaltgrupdurum = "UPDATE bomlistesi SET [Durum] = 'Montajlı', [Tam Op] = [Tam Op] + 1 WHERE [Montaj Durum] = '" + metroGrid5.Rows[row].Cells["dataGridViewButtonColumn4"].Value.ToString() + "' ";
                    SQLiteCommand cmdaltgrupdurum = new SQLiteCommand(sqlaltgrupdurum, baglan);
                    cmdaltgrupdurum.ExecuteNonQuery();
                }
                else
                {
                    sqlmontajdurum = "UPDATE bomlistesi SET [Durum] = 'Montajlı', [Montaj Durum] = [Ust Grup] , [Tam Op] = [Tam Op] + 1 WHERE [ID] = '" + parcaid + "' ";
                    MessageBox.Show(ustgrup);
                }
                SQLiteCommand cmdmontajdurum = new SQLiteCommand(sqlmontajdurum, baglan);
                cmdmontajdurum.ExecuteNonQuery();
                //SQLiteCommand cmd3 = new SQLiteCommand(sql3, baglan);
                //cmd3.ExecuteNonQuery();
                baglan.Close();
            }
            else
            {
                MessageBox.Show("Montaja hazır olmayan parçanın montajı yapılamaz.");
            }
            urunagacimontaj();
            montajrenk();
        }

        private void metroTextBox7_TextChanged(object sender, EventArgs e)
        {
            if (metroTextBox7.Text.ToString() != "")
            {
                string ara = metroTextBox7.Text.ToString();
                SQLiteConnection baglan = new SQLiteConnection(baglanti.database);
                baglan.Open();
                string sqlmontaj = "SELECT [ID], [Grup Adı], [Parça No], [Parça Adı], [Toplam Adet], [Kaynak Durum], [Talaşlı Durum], [Montaj Durum], [Fason], [Lazer], [Durum] FROM bomlistesi WHERE [Montaj Durum] = 'Devam Ediyor' OR [Montaj Durum] = 'Montaj Bekleniyor' AND ([Proje Numarası] LIKE '%" + ara + "%' OR [Parça No] LIKE '%" + ara + "%' OR [Parça Adı] LIKE '%" + ara + "%')";
                damontaj = new SQLiteDataAdapter(sqlmontaj, baglan);
                dtmontaj = new DataTable();
                damontaj.Fill(dtmontaj);
                metroGrid5.DataSource = dtmontaj;
                baglan.Close();
                montajrenk();
            }
            else
            {
                montajparcalari();
                montajrenk();
            }
        }

        private void metroGrid6_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.ColumnIndex == metroGrid6.Columns["dataGridViewButtonColumn5"].Index)
            {
                try
                {
                    string numara = metroGrid6.Rows[e.RowIndex].Cells["dataGridViewButtonColumn5"].Value.ToString();
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

        private void metroLabel1_Click(object sender, EventArgs e)
        {

        }

        private void metroTextBox3_Click(object sender, EventArgs e)
        {

        }


        private void metroTextBox5_TextChanged(object sender, EventArgs e)
        {
            if (metroTextBox5.Text.ToString() != "")
            {
                string ara = metroTextBox5.Text.ToString();
                SQLiteConnection baglan = new SQLiteConnection(baglanti.database);
                baglan.Open();
                string sql = "SELECT [ID], [Grup Adı], [Parça No], [Parça Adı], [Toplam Adet], [Kaynak Durum], [Talaşlı Durum], [Montaj Durum], [Fason], [Lazer], [Durum] FROM bomlistesi WHERE ([Talaşlı Durum] = 'Devam Ediyor' OR [Talaşlı Durum] = 'Fason Talep Ediyor') AND ([Proje Numarası] LIKE '%" + ara + "%' OR [Parça No] LIKE '%" + ara + "%' OR [Parça Adı] LIKE '%" + ara + "%')";
                datalasli = new SQLiteDataAdapter(sql, baglan);
                dttalasli = new DataTable();
                datalasli.Fill(dttalasli);
                metroGrid3.DataSource = dttalasli;
                baglan.Close();
            }
            else
            {
                urunagacitalasli();
            }
        }
        private void metroGrid2_RowValidated(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                SQLiteConnection baglan = new SQLiteConnection(baglanti.database);
                try
                {
                    SQLiteCommandBuilder cmdbldr = new SQLiteCommandBuilder(da);
                    da.Update(dtproje);
                }
                catch (Exception ex) { MessageBox.Show(ex.Message); }
                baglan.Open();
                //KAYNAKLI EKLE   
                string sqldurum = "UPDATE bomlistesi SET [Kaynak Durum] = 'Devam Ediyor' , [Durum] = 'İmalatta' , [Op Sayısı] = [Op Sayısı] + 1 WHERE ([Kaynak Durum] is NULL OR [Kaynak Durum] = '')  AND [Kaynaklı İmalat] = 'VAR' ";
                SQLiteCommand cmddurum = new SQLiteCommand(sqldurum, baglan);
                cmddurum.ExecuteNonQuery();
                //TALAŞLI EKLE
                string sqldurum2 = "UPDATE bomlistesi SET [Talaşlı Durum] = 'Devam Ediyor' , [Durum] = 'İmalatta' , [Op Sayısı] = [Op Sayısı] + 1 WHERE ([Talaşlı Durum] is NULL OR [Talaşlı Durum] = '') AND [Talaşlı İmalat] = 'VAR' ";
                SQLiteCommand cmddurum2 = new SQLiteCommand(sqldurum2, baglan);
                cmddurum2.ExecuteNonQuery();
                //MONTAJ EKLE
                string sqldurum3 = "UPDATE bomlistesi SET [Montaj Durum] = 'Montaj Bekleniyor' , [Durum] = 'Montaj Bekleniyor' , [Op Sayısı] = [Op Sayısı] + 1 WHERE ([Montaj Durum] is NULL OR [Montaj Durum] = '') AND [Montaj] = 'VAR' ";
                SQLiteCommand cmddurum3 = new SQLiteCommand(sqldurum3, baglan);
                cmddurum3.ExecuteNonQuery();
                //UST GRUP EKLE
                int row = metroGrid2.CurrentCell.RowIndex;
                string parcanoustgrup = metroGrid2.Rows[row].Cells["Parcano2"].Value.ToString();
                string ustgrup = parcanoustgrup.Remove(parcanoustgrup.Length - 2, 2) + "00";
                //
                string sqldurum4 = "UPDATE bomlistesi SET [Montaj Durum] = 'Montaj Bekleniyor' , [Ust Grup] = '"+ustgrup+"' , [Op Sayısı] = [Op Sayısı] + 1 WHERE ([Montaj Durum] is NULL OR [Montaj Durum] = '') AND [Montaj2] = 'VAR' ";
                SQLiteCommand cmddurum4 = new SQLiteCommand(sqldurum4, baglan);
                cmddurum4.ExecuteNonQuery();
                baglan.Close();
            }
            catch (Exception ex) { MessageBox.Show(ex.Message); }
        }
        private void metroTabControl1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (metroTabControl1.SelectedTab == metroTabControl1.TabPages["MetroTabPage3"])//your specific tabname
            {
                urunagacikaynakli();
            }
            if (metroTabControl1.SelectedTab == metroTabControl1.TabPages["MetroTabPage4"])//your specific tabname
            {
                urunagacitalasli();
            }
            if (metroTabControl1.SelectedTab == metroTabControl1.TabPages["MetroTabPage5"])//your specific tabname
            {
                fasontalepleri();
            }
            if (metroTabControl1.SelectedTab == metroTabControl1.TabPages["MetroTabPage2"])//your specific tabname
            {
                urunagaci();
            }
            if (metroTabControl1.SelectedTab == metroTabControl1.TabPages["MetroTabPage6"])//your specific tabname
            {
                montajparcalari();
                montajrenk();
            }
            if (metroTabControl1.SelectedTab == metroTabControl1.TabPages["MetroTabPage8"])//your specific tabname
            {
                fasonimalatgecmisi();
            }
        }
        private void montajparcalari()
        {
            SQLiteConnection baglan = new SQLiteConnection(baglanti.database);
            baglan.Open();
            string sqlmontaj = "SELECT [ID], [Grup Adı], [Parça No], [Parça Adı], [Toplam Adet], [Kaynak Durum], [Talaşlı Durum], [Montaj Durum], [Fason], [Lazer], [Durum] FROM bomlistesi WHERE [Montaj Durum] = 'Devam Ediyor' OR [Montaj Durum] = 'Montaj Bekleniyor' ";
            damontaj = new SQLiteDataAdapter(sqlmontaj, baglan);
            dtmontaj = new DataTable();
            damontaj.Fill(dtmontaj);
            metroGrid5.DataSource = dtmontaj;
            baglan.Close();
        }
        private void fasontalepleri()
        {
            SQLiteConnection baglan = new SQLiteConnection(baglanti.database);
            baglan.Open();
            string sql = "SELECT [ID], [Grup Adı], [Parça No], [Parça Adı], [Toplam Adet], [Kaynak Durum], [Talaşlı Durum], [Montaj Durum], [Fason], [Lazer],[Fason Miktarı], [Birim Fiyat], [Toplam], [Durum] FROM bomlistesi WHERE [Kaynak Durum] = 'Fason Talep Ediyor' OR [Talaşlı Durum] = 'Fason Talep Ediyor'  ";
            dafason = new SQLiteDataAdapter(sql, baglan);
            dtfason = new DataTable();
            dafason.Fill(dtfason);
            metroGrid4.DataSource = dtfason;
            baglan.Close();
        }
        private void metroComboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (metroComboBox1.Text.ToString() == "Tümü")
            {
                urunagaci();
                metroTextBox4.Text = "";
            }
            else
            {
                string ara = metroComboBox1.Text.ToString();
                SQLiteConnection baglan = new SQLiteConnection(baglanti.database);
                baglan.Open();
                string sql = "SELECT [ID], [Grup Adı], [Parça No], [Parça Adı], [Toplam Adet], [Kaynaklı İmalat], [Talaşlı İmalat], [Montaj], [Fason], [Lazer], [Birim Fiyat], [Toplam], [Durum] FROM bomlistesi WHERE [Grup Adı] LIKE '%" + ara + "%' ";
                da = new SQLiteDataAdapter(sql, baglan);
                dtproje = new DataTable();
                da.Fill(dtproje);
                metroGrid2.DataSource = dtproje;
                baglan.Close();
            }
        }

        private void kaydisil_Click(object sender, EventArgs e)
        {

        }

        private void metroTextBox3_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                ara.PerformClick();
            }
        }
        private void projeozetiliste()
        {
            try
            {
                projeno = metroTextBox9.Text;
                SQLiteConnection baglan = new SQLiteConnection(baglanti.database);
                baglan.Open();
                string sql = "SELECT [ID], [Grup Adı], [Parça No], [Parça Adı], [Toplam Adet], [Kaynak Durum], [Talaşlı Durum], [Montaj Durum], [Fason], [Fason Durum], [Lazer], [Birim Fiyat], [Toplam], [Durum] FROM bomlistesi WHERE [Proje Numarası]='" + projeno + "'";
                daozet = new SQLiteDataAdapter(sql, baglan);
                dtozet = new DataTable();
                daozet.Fill(dtozet);
                metroGrid6.DataSource = dtozet;
                ////baslik doldur
                string sqlbaslik = "SELECT [Proje No], [Firma Adı], [Proje Adı] FROM yuklenmisbom WHERE [Proje No] ='" + projeno + "'";
                using (SQLiteDataAdapter a = new SQLiteDataAdapter(sqlbaslik, baglan))
                {
                    DataTable t = new DataTable();
                    a.Fill(t);
                    if (t.Rows.Count > 0)
                    {
                        string firmaadi = t.Rows[0]["Firma Adı"].ToString();
                        string projeadi = t.Rows[0]["Proje Adı"].ToString();
                        //Başlık doldur
                        string baslik = projeno + " - " + firmaadi + " - " + projeadi;
                        metroLabel7.Text = baslik;
                    }
                }
                var ilerlenensql = "SELECT TOTAL([Tam Op]) FROM bomlistesi WHERE [Proje Numarası] ='" + projeno + "'";
                using (var cmd = new SQLiteCommand(ilerlenensql, baglan))
                {
                    ilerlenen = Int32.Parse(cmd.ExecuteScalar().ToString());
                }
                var ilerleneceksql = "SELECT TOTAL([Op Sayısı]) FROM bomlistesi WHERE [Proje Numarası] ='" + projeno + "'";
                using (var cmd = new SQLiteCommand(ilerleneceksql, baglan))
                {
                    ilerlenecek = Int32.Parse(cmd.ExecuteScalar().ToString());
                }
                int ilerlemeyuzdesi = (ilerlenen * 100) / ilerlenecek;
                metroLabel8.Text = "Proje ilerlemesi % " + ilerlemeyuzdesi.ToString();
                metroProgressBar3.Value = ilerlemeyuzdesi;
                //Fason Harcaması
                var fasongideri = "SELECT TOTAL([Toplam]) FROM fasontakibi WHERE [Parça no] LIKE '%" + projeno + "%' ";
                using (var cmdfasongideri = new SQLiteCommand(fasongideri, baglan))
                {
                    metroLabel10.Text = String.Format("Toplam tedarikçi gideri: {0} TL.", cmdfasongideri.ExecuteScalar().ToString());
                }
                baglan.Close();
                //saat = 0;

                //FileInfo newFile = new FileInfo(baglanti.calismasaati);
                //ExcelPackage pck = new ExcelPackage(newFile.OpenRead());
                //var ws = pck.Workbook.Worksheets["PROJELISTESI"];
                //string denea = ws.Cells[22, 1].Value.ToString();
                //MessageBox.Show(denea);

                //pck.Dispose();
                //var query1 = (from cell in ws.Cells["a:a"]
                //              where cell.Value.ToString().Equals(projeno)
                //              select cell);
                //foreach (var cell in query1)
                //{
                //    string row = cell.Start.Row.ToString();
                //    saat = Int32.Parse(ws.Cells[cell.Start.Row, 5].Value.ToString()) + saat;
                //}
                //MessageBox.Show(saat.ToString());
            }
            catch { }
        }
        private void projeozetirenk()
        {
            foreach (DataGridViewRow row in metroGrid6.Rows)
            {
                string a = row.Cells["Column14"].Value.ToString();
                string b = row.Cells["Column15"].Value.ToString();
                string c = row.Cells["Column16"].Value.ToString();
                string d = row.Cells["Column17"].Value.ToString();
                string e = row.Cells["Column18"].Value.ToString();
                if ((a == "") && (b == "") && (c == "") && (d == ""))
                {
                    row.DefaultCellStyle.BackColor = Color.LightSalmon;
                    row.DefaultCellStyle.ForeColor = Color.Black;
                }
                else if ((a == "" || a == "Bitti") && (b == "" || b == "Bitti") && (c == "" || c == "Bitti") && (d == "" || d == "Bitti"))
                {
                    row.DefaultCellStyle.BackColor = Color.ForestGreen;
                    row.DefaultCellStyle.ForeColor = Color.Black;
                }
                else if ((e == "Montajlı" || e == "Sevke Hazır"))
                {
                    row.DefaultCellStyle.BackColor = Color.ForestGreen;
                    row.DefaultCellStyle.ForeColor = Color.Black;
                }
                else
                {
                    row.DefaultCellStyle.BackColor = Color.Goldenrod;
                    row.DefaultCellStyle.ForeColor = Color.Black;
                }
            }
        }
        private void montajrenk()
        {
            foreach (DataGridViewRow row in metroGrid5.Rows)
            {
                string a = row.Cells["dataGridViewButtonColumn4"].Value.ToString().Substring(row.Cells["dataGridViewButtonColumn4"].Value.ToString().Length - 2);
                if (a == "00") 
                {
                    row.DefaultCellStyle.BackColor = Color.DimGray;
                    row.DefaultCellStyle.ForeColor = Color.Black;
                }
                else if (row.Cells["dataGridViewTextBoxColumn26"].Value.ToString() == "Montaja Hazır")
                {
                    row.DefaultCellStyle.BackColor = Color.LimeGreen;
                    row.DefaultCellStyle.ForeColor = Color.Black;
                }
                else
                {
                    row.DefaultCellStyle.BackColor = Color.LightSalmon;
                    row.DefaultCellStyle.ForeColor = Color.Black;
                }
            }
        }
        private void fasonimalatgecmisi()
        {
            SQLiteConnection baglan = new SQLiteConnection(baglanti.database);
            baglan.Open();
            string sql = "SELECT [fasonid],  [Parça No], [Grup Adı],[Parça Adı], [Sipariş Tarihi], [Termin], [Sipariş Adedi], [Birim Fiyat], [Toplam], [Onaylayan], [Fason], [İşçilik], [Malzeme], [Kaplama] FROM fasontakibi";
            dafasongecmisi = new SQLiteDataAdapter(sql, baglan);
            dtfasongecmisi = new DataTable();
            dafasongecmisi.Fill(dtfasongecmisi);
            metroGrid7.DataSource = dtfasongecmisi;
            baglan.Close();
        }

    }
}
