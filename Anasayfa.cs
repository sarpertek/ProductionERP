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
using System.Data.OleDb;
using OfficeOpenXml;
using System.Diagnostics;
using System.Data.SQLite;
using DGVPrinterHelper;

namespace _2015
{
    public partial class Anasayfa : MetroFramework.Controls.MetroUserControl
    {
        private static Anasayfa _instance;
        public static Anasayfa Instance
        {
            get
            {
                if (_instance == null)
                    _instance = new Anasayfa();
                return _instance;
            }
        }
        public Anasayfa()
        {
            InitializeComponent();
        }
        private void Anasayfa_Load(object sender, EventArgs e)
        {
            acikisemirleri();
        }

        private void acikisemirleri()
        {
            //FİLTRELE
            FileInfo newFile = new FileInfo(baglanti.exceltakip);
            ExcelPackage pck = new ExcelPackage(newFile);
            var ws = pck.Workbook.Worksheets["Sayfa1"];
            DataTable tbl = new DataTable();
            bool hasHeader = true;
            foreach (var firstRowCell in ws.Cells[1, 1, 1, ws.Dimension.End.Column])
            {
                tbl.Columns.Add(hasHeader ? firstRowCell.Text : string.Format("Column {0}", firstRowCell.Start.Column));
            }
            var startRow = hasHeader ? 2 : 1;
            for (var rowNum = startRow; rowNum <= ws.Dimension.End.Row; rowNum++)
            {
                var wsRow = ws.Cells[rowNum, 1, rowNum, ws.Dimension.End.Column];
                var row = tbl.NewRow();
                foreach (var cell in wsRow)
                {
                    row[cell.Start.Column - 1] = cell.Text;
                }
                tbl.Rows.Add(row);
            }
            String durum = "Malzemeler Hazırlanıyor";
            String durum2 = "Sevk Bekleniyor";
            tbl.DefaultView.RowFilter = "Durum = '" + durum + "' or Durum = '" + durum2 + "'";
            metroGrid1.DataSource = tbl;
            metroComboBox1.Text = "Açık İş Emirleri";
        }

        private void sevkedilmis()
        {
            //FİLTRELE
            FileInfo newFile = new FileInfo(baglanti.exceltakip);
            ExcelPackage pck = new ExcelPackage(newFile);
            var ws = pck.Workbook.Worksheets["Sayfa1"];
            DataTable tbl = new DataTable();
            bool hasHeader = true;
            foreach (var firstRowCell in ws.Cells[1, 1, 1, ws.Dimension.End.Column])
            {
                tbl.Columns.Add(hasHeader ? firstRowCell.Text : string.Format("Column {0}", firstRowCell.Start.Column));
            }
            var startRow = hasHeader ? 2 : 1;
            for (var rowNum = startRow; rowNum <= ws.Dimension.End.Row; rowNum++)
            {
                var wsRow = ws.Cells[rowNum, 1, rowNum, ws.Dimension.End.Column];
                var row = tbl.NewRow();
                foreach (var cell in wsRow)
                {
                    row[cell.Start.Column - 1] = cell.Text;
                }
                tbl.Rows.Add(row);
            }
            String durum = "Sevk Edildi";
            tbl.DefaultView.RowFilter = "Durum = '" + durum + "'";
            metroGrid1.DataSource = tbl;
        }

        private void iptaledilmis()
        {
            //FİLTRELE
            FileInfo newFile = new FileInfo(baglanti.exceltakip);
            ExcelPackage pck = new ExcelPackage(newFile);
            var ws = pck.Workbook.Worksheets["Sayfa1"];
            DataTable tbl = new DataTable();
            bool hasHeader = true;
            foreach (var firstRowCell in ws.Cells[1, 1, 1, ws.Dimension.End.Column])
            {
                tbl.Columns.Add(hasHeader ? firstRowCell.Text : string.Format("Column {0}", firstRowCell.Start.Column));
            }
            var startRow = hasHeader ? 2 : 1;
            for (var rowNum = startRow; rowNum <= ws.Dimension.End.Row; rowNum++)
            {
                var wsRow = ws.Cells[rowNum, 1, rowNum, ws.Dimension.End.Column];
                var row = tbl.NewRow();
                foreach (var cell in wsRow)
                {
                    row[cell.Start.Column - 1] = cell.Text;
                }
                tbl.Rows.Add(row);
            }
            String durum = "İptal Edilmiştir";
            tbl.DefaultView.RowFilter = "Durum = '" + durum + "'";
            metroGrid1.DataSource = tbl;
        }

        private void tumunugoster()
        {
            //Tümünü göster
            FileInfo newFile = new FileInfo(baglanti.exceltakip);
            ExcelPackage pck = new ExcelPackage(newFile);
            var ws = pck.Workbook.Worksheets["Sayfa1"];
            DataTable tbl = new DataTable();
            bool hasHeader = true;
            foreach (var firstRowCell in ws.Cells[1, 1, 1, ws.Dimension.End.Column])
            {
                tbl.Columns.Add(hasHeader ? firstRowCell.Text : string.Format("Column {0}", firstRowCell.Start.Column));
            }
            var startRow = hasHeader ? 2 : 1;
            for (var rowNum = startRow; rowNum <= ws.Dimension.End.Row; rowNum++)
            {
                var wsRow = ws.Cells[rowNum, 1, rowNum, ws.Dimension.End.Column];
                var row = tbl.NewRow();
                foreach (var cell in wsRow)
                {
                    row[cell.Start.Column - 1] = cell.Text;
                }
                tbl.Rows.Add(row);
            }
            metroGrid1.DataSource = tbl;
            //dataGridView1.Columns["Column1"].Visible = true;
        }

        private void metroTile1_Click(object sender, EventArgs e)
        {
            isemriekle isemriekle = new isemriekle();
            isemriekle.ShowDialog();
        }
        private void metroTile3_Click(object sender, EventArgs e)
        {
            Form2 form2 = new Form2();
            form2.ShowDialog();
        }

        private void metroGrid1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.ColumnIndex == 0)
                try
                {
                    FileInfo newFile = new FileInfo(baglanti.exceltakip);
                    ExcelPackage pck = new ExcelPackage(newFile);
                    Detay detay = new Detay();
                    string x0 = metroGrid1.Rows[e.RowIndex].Cells[0].Value.ToString();
                    var ws = pck.Workbook.Worksheets["Sayfa1"];
                    string boxno = "1";
                    var query1 = (from cell in ws.Cells["a:a"]
                                  where cell.Value.ToString().Equals(x0) && ws.Cells[cell.Start.Row, 7].Value.ToString().Equals(boxno)
                                  select cell);
                    foreach (var cell in query1)
                    {
                        string row = cell.Start.Row.ToString();
                        detay.passvalue1 = x0;
                        detay.passvalue2 = ws.Cells[cell.Start.Row, 2].Value.ToString();
                        detay.passvalue3 = ws.Cells[cell.Start.Row, 3].Value.ToString();
                        detay.passvalue4 = ws.Cells[cell.Start.Row, 4].Value.ToString();
                        detay.passvalue5 = ws.Cells[cell.Start.Row, 5].Value.ToString();
                        detay.passvalue6 = ws.Cells[cell.Start.Row, 6].Value.ToString();
                        detay.passvalue7 = ws.Cells[cell.Start.Row, 10].Value.ToString();
                        try
                        {
                            detay.passvalue8 = ws.Cells[cell.Start.Row, 11].Value.ToString();
                        }
                        catch
                        {

                        }
                        detay.passvalue9 = ws.Cells[cell.Start.Row, 12].Value.ToString();
                    }
                    detay.ShowDialog();
                }
                catch { }
            else
                return;
        }
        private void metroTile4_Click(object sender, EventArgs e)
        {
            DialogResult dialogResult = MessageBox.Show("Malzemelerin gönderildiğinden emin misiniz?", "Uyarı", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
            if (dialogResult == DialogResult.Yes)
            {
                try
                {
                    var row = metroGrid1.CurrentCell.RowIndex;
                    string isemrit = metroGrid1.Rows[row].Cells[0].Value.ToString();
                    string box = metroGrid1.Rows[row].Cells[6].Value.ToString();
                    FileInfo newFile = new FileInfo(baglanti.exceltakip);
                    ExcelPackage pck = new ExcelPackage(newFile);
                    var ws = pck.Workbook.Worksheets["Sayfa1"];
                    var query2 = (from cell in ws.Cells["a:a"]
                                  where cell.Value.ToString().Equals(isemrit) && ws.Cells[cell.Start.Row, 7].Value.ToString() == box
                                  select cell);
                    foreach (var cell in query2)
                    {
                        ws.Cells[cell.Start.Row, 10].Value = "Sevk Edildi";
                        ws.Cells[cell.Start.Row, 11].Value = DateTime.Now.ToString("dd/MM/yyyy");
                    }
                    pck.Save();
                    MessageBox.Show(isemrit + " numaralı iş emri başarıyla sevk edildi olarak etiketlenmiştir.");
                }
                catch { MessageBox.Show("Bir hata oluştu."); }
            }
            acikisemirleri();
        }

        private void metroTile5_Click(object sender, EventArgs e)
        {
            DialogResult dialogResult = MessageBox.Show("İş emrini iptal etmek istediğinizden emin misiniz?", "Uyarı", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
            if (dialogResult == DialogResult.Yes)
            {
                try
                {
                    var row = metroGrid1.CurrentCell.RowIndex;
                    string isemrit = metroGrid1.Rows[row].Cells[0].Value.ToString();
                    string box = metroGrid1.Rows[row].Cells[6].Value.ToString();
                    FileInfo newFile = new FileInfo(baglanti.exceltakip);
                    ExcelPackage pck = new ExcelPackage(newFile);
                    var ws = pck.Workbook.Worksheets["Sayfa1"];
                    var query2 = (from cell in ws.Cells["a:a"]
                                  where cell.Value.ToString().Equals(isemrit) && ws.Cells[cell.Start.Row, 7].Value.ToString() == box
                                  select cell);
                    foreach (var cell in query2)
                    {
                        ws.Cells[cell.Start.Row, 10].Value = "İptal Edilmiştir";
                    }
                    pck.Save();
                    MessageBox.Show(isemrit + " numaralı iş emri başarıyla iptal edilmiştir.");
                }
                catch { MessageBox.Show("Bir hata oluştu."); }
            }
            acikisemirleri();
        }

        private void metroComboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                if (metroComboBox1.Text == "Açık İş Emirleri")
                { acikisemirleri(); }
                if (metroComboBox1.Text == "Kapatılmış İş Emirleri")
                { sevkedilmis(); }
                if (metroComboBox1.Text == "İptal Edilmiş İş Emirleri")
                { iptaledilmis(); }
                if (metroComboBox1.Text == "Tüm İş Emirleri")
                { tumunugoster(); }
            }
            catch
            {
                acikisemirleri();
            }
        }

        private void metroTextBox1_TextChanged(object sender, EventArgs e)
        {
            if (metroTextBox1.Text.ToString() != "")
            {
                FileInfo newFile = new FileInfo(baglanti.exceltakip);
                ExcelPackage pck = new ExcelPackage(newFile);
                var ws = pck.Workbook.Worksheets["Sayfa1"];
                DataTable tbl = new DataTable();
                bool hasHeader = true;
                foreach (var firstRowCell in ws.Cells[1, 1, 1, ws.Dimension.End.Column])
                {
                    tbl.Columns.Add(hasHeader ? firstRowCell.Text : string.Format("Column {0}", firstRowCell.Start.Column));
                }
                var startRow = hasHeader ? 2 : 1;
                for (var rowNum = startRow; rowNum <= ws.Dimension.End.Row; rowNum++)
                {
                    var wsRow = ws.Cells[rowNum, 1, rowNum, ws.Dimension.End.Column];
                    var row = tbl.NewRow();
                    foreach (var cell in wsRow)
                    {
                        row[cell.Start.Column - 1] = cell.Text;
                    }
                    tbl.Rows.Add(row);
                    DataView datafilter = tbl.DefaultView;
                    string ara = metroTextBox1.Text.ToString();
                    datafilter.RowFilter = "[Durum] LIKE '%"+ara+ "%' OR [İş Emri No] LIKE '%" + ara + "%' OR [UC İş Emri] LIKE '%" + ara + "%' OR [Müşteri Adı] LIKE '%" + ara + "%' OR [Bölge] LIKE '%" + ara + "%' ";
                    //OR Detay LIKE '"+ara+"'
                }
                metroComboBox1.PromptText = "Arama Sonuçları";
                metroGrid1.DataSource = tbl;
            }
            else { acikisemirleri(); }

            
        }

        private void metroTextBox1_Click(object sender, EventArgs e)
        {

        }

        private void metroTile2_Click(object sender, EventArgs e)
        {
            Duzenle duzenle = new Duzenle();
            duzenle.ShowDialog();
        }

        private void metroGrid1_Click(object sender, EventArgs e)
        {
        }

        private void metroTile6_Click(object sender, EventArgs e)
        {
            if (metroTile1.Enabled == false)
            {
                anasayfa2 anasayfa2 = new anasayfa2();
                anasayfa2.metroTile2.Enabled = false;
                anasayfa2.Show();
                this.Hide();
            }
            else
            {
                anasayfa2 anasayfa2 = new anasayfa2();
                anasayfa2.Show();
                this.Hide();
            }
        }

        private void metroButton1_Click(object sender, EventArgs e)
        {
            DGVPrinter printer = new DGVPrinter();
            printer.KeepRowsTogether = true;
            printer.Title = metroComboBox1.PromptText;
            printer.SubTitle = "Tarih: " + DateTime.Now.ToString("dd/MM/yyyy");
            printer.SubTitleFormatFlags = StringFormatFlags.LineLimit | StringFormatFlags.NoClip;
            printer.PageNumbers = true;
            printer.PageNumberInHeader = false;
            printer.PorportionalColumns = true;
            printer.HeaderCellAlignment = StringAlignment.Near;
            //printer.Footer = "Taban";
            printer.FooterSpacing = 15;
            printer.FooterSpacing = 15;
            metroGrid1.Theme = MetroFramework.MetroThemeStyle.Light;
            metroGrid1.Style = MetroFramework.MetroColorStyle.Black;
            printer.PrintDataGridView(metroGrid1);
            metroGrid1.Theme = MetroFramework.MetroThemeStyle.Dark;
            metroGrid1.Style = MetroFramework.MetroColorStyle.Orange;
        }

        private void metroTile6_Click_1(object sender, EventArgs e)
        {
            DialogResult dialogResult = MessageBox.Show("Faturaların kesildiğinden emin misiniz?", "Uyarı", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
            if (dialogResult == DialogResult.Yes)
            {
                try
                {
                    var row = metroGrid1.CurrentCell.RowIndex;
                    string isemrit = metroGrid1.Rows[row].Cells[0].Value.ToString();
                    string box = metroGrid1.Rows[row].Cells[6].Value.ToString();
                    FileInfo newFile = new FileInfo(baglanti.exceltakip);
                    ExcelPackage pck = new ExcelPackage(newFile);
                    var ws = pck.Workbook.Worksheets["Sayfa1"];
                    var query2 = (from cell in ws.Cells["a:a"]
                                  where cell.Value.ToString().Equals(isemrit) && ws.Cells[cell.Start.Row, 7].Value.ToString() == box
                                  select cell);
                    foreach (var cell in query2)
                    {
                        ws.Cells[cell.Start.Row, 13].Value = "Kesildi";
                    }
                    pck.Save();
                    MessageBox.Show(isemrit + " numaralı iş emrinin faturası başarı ile eklendi olarak kaydedilmiştir.");
                }
                catch { MessageBox.Show("Bir hata oluştu."); }
            }
            acikisemirleri();
        }

        private void metroTile2_Click_1(object sender, EventArgs e)
        {
            acikisemirleri();
        }
    }
}
