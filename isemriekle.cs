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
using Outlook = Microsoft.Office.Interop.Outlook;
using System.Text.RegularExpressions;

namespace _2015
{
    public partial class isemriekle : MetroForm
    {
        string usrnme = Environment.UserName;
        public isemriekle()
        {
            InitializeComponent();
            dateTimePicker1.MinDate = DateTime.Now;
        }

        private void veriekle()
        {
            FileInfo newFile = new FileInfo(baglanti.exceltakip);
            ExcelPackage pck = new ExcelPackage(newFile);
            //Add the Content sheet
            var ws = pck.Workbook.Worksheets["Sayfa1"];
            ws.View.ShowGridLines = false;
            //Başlık
            ws.Cells["A1"].Value = "İş Emri No";
            ws.Cells["B1"].Value = "UC İş Emri";
            ws.Cells["C1"].Value = "Müşteri Adı";
            ws.Cells["D1"].Value = "Giriş Tarihi";
            ws.Cells["E1"].Value = "Termin Tarihi";
            ws.Cells["F1"].Value = "Bölge";
            ws.Cells["G1"].Value = "Box No";
            ws.Cells["H1"].Value = "Box Bilgileri";
            ws.Cells["I1"].Value = "Box Tarihi";
            ws.Cells["J1"].Value = "Durum";
            ws.Cells["K1"].Value = "Sevk Tarihi";
            ws.Cells["L1"].Value = "Notlar";
            ws.Cells["M1"].Value = "Fatura Durumu";
            ws.Cells["A1:M1"].Style.Font.Bold = true;
            ws.Cells["A1:M1"].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            ws.Cells["A1:M1"].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Black);
            ws.Cells["A1:M1"].Style.Font.Color.SetColor(Color.White);

            int sonrow = ws.Dimension.End.Row + 1;
            string isemrit = String.Concat(metroComboBox1.Text, "/", metroComboBox2.Text, "/", metroTextBox1.Text);
            if (metroComboBox3.Text == "" || metroTextBox2.Text == "")
            {
                string ucisemrit = "Belirtilmemiş";
                ws.Cells[sonrow, 2].Value = ucisemrit;
            }
            if (metroComboBox3.Text != "" && metroTextBox2.Text != "")
            {
                string ucisemrit = String.Concat(metroComboBox3.Text, "/UC/", metroTextBox2.Text);
                ws.Cells[sonrow, 2].Value = ucisemrit;
            }
            string musteri = musteritext.Text.ToString();
            string projeadi = projetext.Text.ToString();
            string notlar = notlartext.Text.ToString();
            string termin = dateTimePicker1.Value.ToString("dd/MM/yyyy");
            ws.Cells[sonrow, 1].Value = isemrit;
            ws.Cells[sonrow, 3].Value = musteri;
            ws.Cells[sonrow, 4].Value = DateTime.Now.ToString("dd/MM/yyyy");
            ws.Cells[sonrow, 5].Value = termin;
            ws.Cells[sonrow, 6].Value = projeadi;
            ws.Cells[sonrow, 7].Value = "1";
            ws.Cells[sonrow, 10].Value = "Malzemeler hazırlanıyor";
            ws.Cells[sonrow, 12].Value = notlar;
            ws.Cells[sonrow, 13].Value = "Kesilmedi";
            //Create an autofilter for the range
            ws.Cells["A1:M1"].AutoFilter = true;
            ws.Cells.AutoFitColumns(0);  //Autofit columns for all cells
            pck.Save();
        }

        private void dosyaolustur()
        {
            DataTable dt = new DataTable();
            foreach (DataGridViewColumn col in dataGridView1.Columns)
            {
                dt.Columns.Add(col.HeaderText);
            }

            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                DataRow dRow = dt.NewRow();
                foreach (DataGridViewCell cell in row.Cells)
                {
                    dRow[cell.ColumnIndex] = cell.Value;
                }
                dt.Rows.Add(dRow);
            }
            string isemrid = String.Concat(metroComboBox1.Text, "-" , metroComboBox2.Text, "-", metroTextBox1.Text);
            if (metroComboBox3.Text != "" || metroTextBox2.Text != "")
            {
                string ucisemrit = String.Concat(metroComboBox3.Text, "/UC/", metroTextBox2.Text);
                string musteri = musteritext.Text.ToString();
                string projeadi = projetext.Text.ToString();
                string notlar = notlartext.Text.ToString();
                string termin = dateTimePicker1.Value.ToString("dd/MM/yyyy");
                string bugun = DateTime.Now.ToString("dd/MM/yyyy");
                FileInfo newFile = new FileInfo("Z:\\TM-KLT TAKIP\\Malzeme listeleri\\" + isemrid + " " + musteri + "-" + projeadi + ".xlsx");
                ExcelPackage pck1 = new ExcelPackage(newFile);
                var ws = pck1.Workbook.Worksheets.Add("Sayfa1");
                ws.Cells["A1"].LoadFromDataTable(dt, true);
                //ws.Cells["B1"].LoadFromDataTable(dt2, true);
                ws.Cells["C1"].Value = "Sandık No";
                ws.Cells["A1:C1"].Style.Font.Bold = true;
                ws.Cells["A1:C1"].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                ws.Cells["A1:C1"].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Black);
                ws.Cells["A1:C1"].Style.Font.Color.SetColor(Color.White);
                ws.Cells["A1:B1"].AutoFilter = true;
                ws.Cells.AutoFitColumns(0);  //Autofit columns for all cells
                pck1.Save();
            }
            else
            {
                string musteri = musteritext.Text.ToString();
                string projeadi = projetext.Text.ToString();
                string notlar = notlartext.Text.ToString();
                string termin = dateTimePicker1.Value.ToString("dd/MM/yyyy");
                string bugun = DateTime.Now.ToString("dd/MM/yyyy");
                FileInfo newFile = new FileInfo("Z:\\TM-KLT TAKIP\\Malzeme listeleri\\" + isemrid + " " + musteri + "-" + projeadi + ".xlsx");
                ExcelPackage pck1 = new ExcelPackage(newFile);
                var ws = pck1.Workbook.Worksheets.Add("Sayfa1");
                ws.Cells["A1"].LoadFromDataTable(dt, true);
                //ws.Cells["B1"].LoadFromDataTable(dt2, true);
                ws.Cells["C1"].Value = "Sandık No";
                ws.Cells["A1:C1"].Style.Font.Bold = true;
                ws.Cells["A1:C1"].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                ws.Cells["A1:C1"].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Black);
                ws.Cells["A1:C1"].Style.Font.Color.SetColor(Color.White);

                ws.Cells["A1:B1"].AutoFilter = true;
                ws.Cells.AutoFitColumns(0);  //Autofit columns for all cells
                pck1.Save();
            }
        }


        private void isemriekle_Load(object sender, EventArgs e)
        {
        }

        private void metroButton1_Click(object sender, EventArgs e)
        {
            veriekle();
            dosyaolustur();

            string isemrid = String.Concat(metroComboBox1.Text, "-", metroComboBox2.Text, "-", metroTextBox1.Text);
            if (metroComboBox3 != null || metroTextBox2.Text != null)
            {
                string ucisemrit = String.Concat(metroComboBox3.Text, "/UC/", metroTextBox2.Text);
                string isemrit = String.Concat(metroComboBox1.Text, "/", metroComboBox2.Text, "/", metroTextBox1.Text);
                string musteri = musteritext.Text.ToString();
                string projeadi = projetext.Text.ToString();
                string notlar = notlartext.Text.ToString();
                string termin = dateTimePicker1.Value.ToString("dd/MM/yyyy");
                string bugun = DateTime.Now.ToString("dd/MM/yyyy");
                string filesw = "Z:\\TM-KLT TAKIP\\Malzeme listeleri\\" + isemrid + " " + musteri + "-" + projeadi + ".xlsx";
                // Create the Outlook application.
                Outlook.Application oApp = new Outlook.Application();
                // Create a new mail item.
                Outlook.MailItem oMsg = (Outlook.MailItem)oApp.CreateItem(Outlook.OlItemType.olMailItem);
                // Set HTMLBody. 
                //add the body of the email
                oMsg.HTMLBody = "İş emri malzeme listesi linki:  " + " <a href= " + '\u0022' + filesw + '\u0022' + ">" + isemrit + "</a>" + "<br>" + "<br>" + "Lütfen gönderilecek malzeme listesini yalnızca yukarıdaki linkten güncelleyiniz." + "<br>" + "<br>" +
                                "<b>İş emri numarası: </b>" + isemrit + "<br>" +
                                "<b>Müşteri:          </b>" + musteri + "<br>" +
                                "<b>UC iş emri no:    </b>" + ucisemrit + "<br>" +
                                "<b>Giriş tarihi:     </b>" + bugun + "<br>" +
                                "<b>Tahmini paketleme tarihi:     </b>" + termin + "<br>" +
                                "<b>Proje adı:        </b>" + projeadi + "<br>" +
                                "<b>Notlar: </b>" + notlar;
                //Subject line
                oMsg.Subject = isemrit + " " + musteri + " " + "İş Emri Açılmıştır";
                // Add a recipient.
                Outlook.Recipients oRecips = (Outlook.Recipients)oMsg.Recipients;
                // Change the recipient in the next line if necessary.
                //Outlook.Recipient oRecip = (Outlook.Recipient)oRecips.Add("sarper.tek@unimak.com");
                //oRecip.Resolve();
                // Send.
                oMsg.Display();
                // Clean up.
                //oRecip = null;
                oRecips = null;
                oMsg = null;
                oApp = null;
                this.Close();
            }
            else
            {
                string ucisemrit = "Belirtilmemiş";
                string isemrit = String.Concat(metroComboBox1.Text, "/", metroComboBox2.Text, "/", metroTextBox1.Text);
                string musteri = musteritext.Text.ToString();
                string projeadi = projetext.Text.ToString();
                string notlar = notlartext.Text.ToString();
                string termin = dateTimePicker1.Value.ToString("dd/MM/yyyy");
                string bugun = DateTime.Now.ToString("dd/MM/yyyy");
                string filesw = "Z:\\TM-KLT TAKIP\\Malzeme listeleri\\" + isemrid + " " + musteri + "-" + projeadi + ".xlsx";
                // Create the Outlook application.
                Outlook.Application oApp = new Outlook.Application();
                // Create a new mail item.
                Outlook.MailItem oMsg = (Outlook.MailItem)oApp.CreateItem(Outlook.OlItemType.olMailItem);
                // Set HTMLBody. 
                //add the body of the email
                oMsg.HTMLBody = "İş emri malzeme listesi linki:  " + " <a href= " + '\u0022' + filesw + '\u0022' + ">" + isemrit + "</a>" + "<br>" + "<br>" + "Lütfen gönderilecek malzeme listesini yalnızca yukarıdaki linkten güncelleyiniz." + "<br>" + "<br>" +
                                "<b>İş emri numarası: </b>" + isemrit + "<br>" +
                                "<b>Müşteri:          </b>" + musteri + "<br>" +
                                "<b>UC iş emri no:    </b>" + ucisemrit + "<br>" +
                                "<b>Giriş tarihi:     </b>" + bugun + "<br>" +
                                "<b>Tahmini paketleme tarihi:     </b>" + termin + "<br>" +
                                "<b>Ülke/Şehir:        </b>" + projeadi + "<br>" +
                                "<b>Notlar: </b>" + notlar;
                //Subject line
                oMsg.Subject = isemrit + " " + musteri + " " + "İş Emri Açılmıştır";
                // Add a recipient.
                Outlook.Recipients oRecips = (Outlook.Recipients)oMsg.Recipients;
                // Change the recipient in the next line if necessary.
                //Outlook.Recipient oRecip = (Outlook.Recipient)oRecips.Add("sarper.tek@unimak.com");
                //oRecip.Resolve();
                // Send.
                oMsg.Display();
                // Clean up.
                //oRecip = null;
                oRecips = null;
                oMsg = null;
                oApp = null;
                this.Close();
            }
        }

        private void metroButton2_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void yapıştırToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {

            DataObject o = (DataObject)Clipboard.GetDataObject();
                if (o.GetDataPresent(DataFormats.Text))
                {
                    //if (dataGridView1.RowCount > 0)
                    //    dataGridView1.Rows.Clear();
                    string[] pastedRows = Regex.Split(o.GetData(DataFormats.Text).ToString().TrimEnd("\r\n".ToCharArray()), "\r\n");
                    int j = dataGridView1.CurrentCell.RowIndex;
                    int k = dataGridView1.CurrentCell.ColumnIndex;
                    foreach (string pastedRow in pastedRows)
                    {
                        string[] pastedRowCells = pastedRow.Split(new char[] { '\t' });

                        dataGridView1.Rows.Add();
                        //int myRowIndex = dataGridView1.Rows.Count - 1;

                        using (DataGridViewRow myDataGridViewRow = dataGridView1.Rows[j])
                        {
                            for (int i = 0; i < pastedRowCells.Length; i++)
                            {
                                try
                                {
                                    int l = k + i;
                                    myDataGridViewRow.Cells[l].Value = pastedRowCells[i];
                                }
                                catch { }
                            }
                        }
                        j++;
                    }
                }
            }
            catch { }
        }
    }
}

