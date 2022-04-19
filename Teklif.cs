using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml;
using OfficeOpenXml;
using System.IO;

namespace _2015
{
    public partial class Teklif : MetroFramework.Controls.MetroUserControl
    {
        int baslangic { get; set; }
        string offernumber { get; set; }
        private static Teklif _instance;
        public static Teklif Instance
        {
            get
            {
                if (_instance == null)
                    _instance = new Teklif();
                return _instance;
            }
        }
        public Teklif()
        {
            InitializeComponent();
        }
        private void kurbelirle()
        {
            try
            {
                // Bugün (en son iş gününe) e ait döviz kurları için
                string today = "http://www.tcmb.gov.tr/kurlar/today.xml";

                // 14 Şubat 2013 e ait döviz kurları için
                //string anyDays = "http://www.tcmb.gov.tr/kurlar/201302/14022013.xml";

                var xmlDoc = new XmlDocument();
                xmlDoc.Load(today);

                // Xml içinden tarihi alma - gerekli olabilir
                DateTime exchangeDate = Convert.ToDateTime(xmlDoc.SelectSingleNode("//Tarih_Date").Attributes["Tarih"].Value);

                string USD = xmlDoc.SelectSingleNode("Tarih_Date/Currency[@Kod='USD']/BanknoteSelling").InnerXml;
                string EURO = xmlDoc.SelectSingleNode("Tarih_Date/Currency[@Kod='EUR']/BanknoteSelling").InnerXml;
                string POUND = xmlDoc.SelectSingleNode("Tarih_Date/Currency[@Kod='GBP']/BanknoteSelling").InnerXml;
                //metroLabel14.Text = EURO;
                //metroLabel15.Text = DateTime.Now.ToString("dd/MM/yyyy");
                metroLabel15.Text = (string.Format("TCMB EURO KURU  (Tarih {0}) 1 Euro = {1}", exchangeDate.ToShortDateString(), EURO, " TL"));
                eurotext.Text = EURO;
            }
            catch
            {
                eurotext.Text = "";
            }
        }

        private void metroButton1_Click(object sender, EventArgs e)
        {
            verikontrol();
        }

        private void metroLabel13_Click(object sender, EventArgs e)
        {

        }

        private void Teklif_Load(object sender, EventArgs e)
        {
            kurbelirle();
        }
        private void dosyaolustur()
        {
            FileInfo newFile = new FileInfo(baglanti.offertemplate);
            ExcelPackage pck = new ExcelPackage(newFile);
            //Add the Content sheet
            var ws = pck.Workbook.Worksheets["Sayfa1"];
            ws.View.ShowGridLines = false;
            //Başlık
            ws.Cells["E24"].Value = ilgilitext.Text;
            ws.Cells["C17"].Value = musteritext.Text;
            ws.Cells["C18:D19"].Merge = true;
            ws.Cells["C18:D19"].Style.WrapText = true;
            ws.Cells["C18"].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Top;
            ws.Cells["C18"].Value = adrestext.Text;
            ws.Cells["C20"].Value = telefontext.Text;
            ws.Cells["C21"].Value = faxtext.Text;
            ws.Cells["K10"].Value = offernumber;
            ws.Cells["K11"].Value = DateTime.Now.ToString("dd/mm/yyyy");
            ws.Cells["K12:L13"].Merge = true;
            ws.Cells["K12:L13"].Style.WrapText = true;
            ws.Cells["K12"].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Top;
            ws.Cells["K12"].Value = paymenttermstext.Text;
            ws.Cells["K15"].Value = deliverytimetext.Text;
            ws.Cells["K16"].Value = deliverytermtext.Text;
            if (transportcheck.Checked == true)
            { ws.Cells["K17"].Value = "INCLUDED"; }
            else
            { ws.Cells["K17"].Value = "EXCLUDED"; }
            if (insurancecheck.Checked == true)
            { ws.Cells["K18"].Value = "INCLUDED"; }
            else
            { ws.Cells["K18"].Value = "EXCLUDED"; }
            if (taxescheck.Checked == true)
            { ws.Cells["K19"].Value = "INCLUDED"; }
            else
            { ws.Cells["K19"].Value = "EXCLUDED"; }
            if (packingcheck.Checked == true)
            { ws.Cells["K20"].Value = "INCLUDED"; }
            else
            { ws.Cells["K20"].Value = "EXCLUDED"; }
            ws.Cells["K21"].Value = Validitytext.Text;

            for (int j=0; j<=dataGridView1.Rows.Count - 2; j++)
                {
                int startrow = 29 + j;
                ws.Cells[29 + j, 2, 29 + j, 7].Merge = true;
                ws.Cells[29 + j, 2].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
                ws.Cells[29 + j, 8].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                ws.Cells[29 + j, 9, 29 + j, 10].Merge = true;
                ws.Cells[29 + j, 9].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Right;
                ws.Cells[29 + j, 11, 29 + j, 13].Merge = true;
                ws.Cells[29 + j, 11].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Right;
                ws.Cells[startrow, 2].Value = dataGridView1.Rows[j].Cells[Description.Index].Value.ToString();
                ws.Cells[startrow, 8].Value = dataGridView1.Rows[j].Cells[Qty.Index].Value.ToString();
                ws.Cells[startrow, 9].Value = dataGridView1.Rows[j].Cells[Unitprice.Index].Value.ToString();
                ws.Cells[startrow, 11].Value = dataGridView1.Rows[j].Cells[TotalPrice.Index].Value.ToString();
                }

            ws.Cells[29 + dataGridView1.Rows.Count - 1, 2, 29 + dataGridView1.Rows.Count - 1, 10].Merge = true;
            ws.Cells[29 + dataGridView1.Rows.Count - 1, 11, 29 + dataGridView1.Rows.Count - 1, 13].Merge = true;
            ws.Cells[29 + dataGridView1.Rows.Count - 1, 2].Value = "OFFER TOTAL PRICE";
            ws.Cells[29 + dataGridView1.Rows.Count - 1, 2].Style.Font.Bold = true;
            ws.Cells[29 + dataGridView1.Rows.Count - 1, 2].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            ws.Cells[29 + dataGridView1.Rows.Count - 1, 2].Style.Fill.BackgroundColor.SetColor(Color.LightGray);
            ws.Cells[29 + dataGridView1.Rows.Count - 1, 11].Value = metroLabel11.Text.ToString();
            ws.Cells[29 + dataGridView1.Rows.Count - 1, 11].Style.Font.Bold = true;
            ws.Cells[29 + dataGridView1.Rows.Count - 1, 11].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            ws.Cells[29 + dataGridView1.Rows.Count - 1, 11].Style.Fill.BackgroundColor.SetColor(Color.LightGray);
            ws.Cells[29 + dataGridView1.Rows.Count - 1, 11].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Right;
            ws.Cells["P20"].Value = eurotext.Text ;
            string path = @"C:\Users\uretim6\Desktop\" + "deneme.xlsx";
            Stream stream = File.Create(path);
            pck.SaveAs(stream);
            stream.Close();
        }
        private void veriekle()
        {
            FileInfo newFile = new FileInfo(baglanti.offerlist);
            ExcelPackage pck = new ExcelPackage(newFile);
            var ws = pck.Workbook.Worksheets.First();
            int sira = Int32.Parse(ws.Cells[ws.Dimension.End.Row , 1].Value.ToString());
            int yenisira = sira + 1;
            //
            int sonrow = ws.Dimension.End.Row + 1;
            ws.Cells[sonrow, 1].Value = yenisira.ToString();
            ws.Cells[sonrow, 2].Value = musteritext.Text;
            // Ülkeyi ayır
            string ulke2 = adrestext.Text.ToString();
            int sonindex = ulke2.Length - 1; // son karakterin yeri
            try
            {
                for (int index = sonindex; index >= 0; --index)
                {
                    if (!Char.IsLetter(ulke2[index]))
                    {
                        MessageBox.Show(index.ToString());
                        baslangic = index + 1;
                        break;
                    }
                }
            }
            catch
            {
            }
            string ulke = ulke2.Substring(baslangic);
            ws.Cells[sonrow, 3].Value = ulke;
            //
            ws.Cells[sonrow, 4].Value = ilgilitext.Text;
            string yil = DateTime.Now.ToString("yyyy");
            offernumber = string.Concat(yil + "/TKF YP/" + yenisira);
            ws.Cells[sonrow, 6].Value = offernumber;
            ws.Cells[sonrow, 7].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Right;
            ws.Cells[sonrow, 7].Value = DateTime.Now.ToString("dd/mm/yyyy");
            ws.Cells[sonrow, 8].Value = "Spare Parts";
            pck.Save();
        }
        private void dataGridView1_CellValidated(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex > -1)
            {
                try
                {
                    DataGridViewRow row = dataGridView1.Rows[e.RowIndex];
                    // . yı sil , yap
                    string valueA = row.Cells[Qty.Index].Value.ToString().Replace('.', ',');
                    string valueB = row.Cells[Unitprice.Index].Value.ToString().Replace('.', ',');
                    string valueD = row.Cells[Maliyetcarpani.Index].Value.ToString().Replace('.', ',');
                    row.Cells[Qty.Index].Value = valueA;
                    row.Cells[Unitprice.Index].Value = valueB;
                    row.Cells[Maliyetcarpani.Index].Value = valueD;
                    float result;
                    if (float.TryParse(valueA, out result)
                        && float.TryParse(valueB, out result)
                        && float.TryParse(valueD, out result))
                    {
                        float valueC = float.Parse(valueA) * float.Parse(valueB) * float.Parse(valueD);
                        row.Cells[TotalPrice.Index].Value = Math.Round(valueC, 2, MidpointRounding.AwayFromZero).ToString();
                        try
                        {
                            float toplam = 0;
                            for (int i=0; i<= dataGridView1.Rows.Count; i++)
                                {
                                string totalvalue = dataGridView1.Rows[i].Cells[TotalPrice.Index].Value.ToString();
                                toplam = float.Parse(totalvalue) + toplam;
                                metroLabel11.Text = Math.Round(toplam, 2, MidpointRounding.AwayFromZero).ToString() + " €"; 
                                }
                        }
                        catch { }
                    }
                }
                catch
                {
                }
            }
        }
        private void verikontrol()
        { 
            if (ilgilitext.Text == "" || musteritext.Text == "" || adrestext.Text =="" || telefontext.Text =="" || faxtext.Text == "" || paymenttermstext.Text=="" || deliverytermtext.Text=="" || deliverytimetext.Text=="" || Validitytext.Text=="" || eurotext.Text=="" || dataGridView1.Rows.Count == 0)
            {
                MessageBox.Show("Lütfen tüm alanları doldurunuz.");
            }
            else
            {
                veriekle();
                dosyaolustur();
            }
        }
    }
}
