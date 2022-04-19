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
using OfficeOpenXml;
using System.IO;
using DGVPrinterHelper;

namespace _2015
{
    public partial class Detay : MetroForm
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

        private void metroButton1_Click(object sender, EventArgs e)
        {
            string[] words = isemri.Split('/');
            string isemrid = string.Join("-", words);
            System.Diagnostics.Process.Start("Z:\\TM-KLT TAKIP\\Malzeme listeleri\\" + isemrid + " " + musteriadi + "-" + projeadi + ".xlsx");
        }

        private void metroButton2_Click(object sender, EventArgs e)
        {
            DGVPrinter printer = new DGVPrinter();
            printer.Title = isemri + " Malzeme listesi";
            printer.SubTitle = "Tarih: " + DateTime.Now.ToString("dd/MM/yyyy");
            printer.SubTitleFormatFlags = StringFormatFlags.LineLimit | StringFormatFlags.NoClip;
            printer.PageNumbers = true;
            printer.PageNumberInHeader = false;
            printer.PorportionalColumns = true;
            printer.HeaderCellAlignment = StringAlignment.Near;
            //printer.Footer = "Taban";
            printer.FooterSpacing = 15;
            metroGrid1.Theme = MetroFramework.MetroThemeStyle.Light;
            metroGrid1.Style = MetroFramework.MetroColorStyle.Black;
            printer.PrintDataGridView(metroGrid1);
            metroGrid1.Theme = MetroFramework.MetroThemeStyle.Dark;
            metroGrid1.Style = MetroFramework.MetroColorStyle.Orange;
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

        public Detay()
        {
            InitializeComponent();
        }

        private void Detay_Load(object sender, EventArgs e)
        {
            label10.Text = isemri;
            label11.Text = ucisemri;
            label12.Text = musteriadi;
            label13.Text = giristarihi;
            label14.Text = paketlemetarihi;
            label15.Text = projeadi;
            label16.Text = durum;
            label17.Text = sevktarihi;
            label18.Text = notlar;
            label18.ScrollBars = ScrollBars.Vertical;
            //Tümünü göster
            try
            {
                string[] words = isemri.Split('/');
            string isemrid = string.Join("-", words);
            string filesw = "Z:\\TM-KLT TAKIP\\Malzeme listeleri\\" + isemrid + " " + musteriadi + "-" + projeadi + ".xlsx";

                FileInfo newFile = new FileInfo(filesw);
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
                metroGrid1.Visible = true;
            }
            catch
            {
                MessageBox.Show("Malzeme listesi başka bir bilgisayarda açık ya da malzeme listesi bulunamadı.");
            }
            return;
        }
    }
}
