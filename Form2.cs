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
using System.IO;
using Microsoft.Office.Core;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace _2015
{
    public partial class Form2 : MetroForm
    {
        public Form2()
        {
            InitializeComponent();
        }

        private void Form2_Load(object sender, EventArgs e)
        {
            acikisemirleri();
            foreach (DataGridViewRow row in metroGrid1.Rows)
            {
                comboBox1.Items.Add(row.Cells[0].Value.ToString());
            }
        }
        private void boxbilgisiekle()
        {
            //ARA
            string usrnme = Environment.UserName;
            String isemrit = comboBox1.Text.ToString();
            FileInfo newFile = new FileInfo(baglanti.exceltakip);
            ExcelPackage pck = new ExcelPackage(newFile);
            var ws = pck.Workbook.Worksheets["Sayfa1"];
            if (checkBox1.Checked == true)
            {
                var query1 = (from cell in ws.Cells["a:a"]
                              where cell.Value.ToString().Equals(isemrit) && ws.Cells[cell.Start.Row, 10].Value.ToString() == "Malzemeler hazırlanıyor"
                              select cell);
                foreach (var cell in query1)
                {
                    int sat = cell.Start.Row;
                    string boxbilgisis;
                    if (textBox1.Text == "" || textBox2.Text == "" || textBox3.Text == "" || textBox4.Text == "")
                    {
                        boxbilgisis = "Kargo";
                        ws.Cells[sat, 8].Value = boxbilgisis;
                    }
                    else
                    {
                        boxbilgisis = String.Concat(textBox1.Text, "x", textBox2.Text, "x", textBox3.Text, "cm", " ", textBox4.Text, "kg");
                        ws.Cells[sat, 8].Value = boxbilgisis;
                    }
                    ws.Cells[sat, 9].Value = DateTime.Now.ToString("dd/MM/yyyy"); ;
                    ws.Cells[sat, 10].Value = "Sevk Bekleniyor";
                    string musteri = ws.Cells[sat, 3].Value.ToString();
                    string projeadi = ws.Cells[sat, 6].Value.ToString();
                    Outlook.Application oApp = new Outlook.Application();
                    string[] x1 = isemrit.Split('/');
                    string isemrid = string.Join("-", x1);
                    string yol = ("Z:\\TM-KLT TAKIP\\Malzeme listeleri\\" + isemrid + " " + musteri + "-" + projeadi + ".xlsx");
                    Outlook.MailItem oMsg = (Outlook.MailItem)oApp.CreateItem(Outlook.OlItemType.olMailItem);
                    oMsg.HTMLBody = "İş emri malzeme listesi:  " + " <a href= " + '\u0022' + yol + '\u0022' + ">" + isemrit + "</a>" + "<br>" + "<br>" + "Lütfen gönderilecek malzeme listesini yalnızca yukarıdaki linkten güncelleyiniz." + "<br>" + "Bu iş emrine başka sandık eklenmeyecektir." + "<br>" + "<br>" +
                                    "<b>İş emri numarası: </b>" + isemrit + "<br>" +
                                    "<b>Sandık boyutları:    </b>" + boxbilgisis + "<br>" +
                                    "<b>Notlar: </b>" + ws.Cells[sat, 12].Value.ToString() ;
                    //Subject line
                    oMsg.Subject = isemrit + " " + musteri + " İş emrine ait sandık hazırlanmıştır";
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
                }
            }
            else
            {
                var query2 = (from cell in ws.Cells["a:a"]
                              where cell.Value.ToString().Equals(isemrit) && ws.Cells[cell.Start.Row, 10].Value.ToString() == "Malzemeler hazırlanıyor"
                              select cell);
                int sonrow = ws.Dimension.End.Row + 1;
                foreach (var cell in query2)
                {
                    int sat = cell.Start.Row;
                    string ucisemrit = ws.Cells[sat, 2].Value.ToString();
                    string musteri = ws.Cells[sat, 3].Value.ToString();
                    string projeadi = ws.Cells[sat, 6].Value.ToString();
                    string notlar = ws.Cells[sat, 12].Value.ToString();
                    string termin = ws.Cells[sat, 5].Value.ToString();
                    string giristarihi = ws.Cells[sat, 4].Value.ToString();
                    string boxno = ws.Cells[sat, 7].Value.ToString();
                    string boxbilgisis;
                    if (textBox1.Text == "" || textBox2.Text == "" || textBox3.Text == "" || textBox4.Text == "")
                    {
                        boxbilgisis = "Kargo";
                        ws.Cells[sat, 8].Value = boxbilgisis;
                    }
                    else
                    {
                        boxbilgisis = String.Concat(textBox1.Text, "x", textBox2.Text, "x", textBox3.Text, "cm", " ", textBox4.Text, "kg");
                        ws.Cells[sat, 8].Value = boxbilgisis;
                    }
                    ws.Cells[sat, 9].Value = DateTime.Now.ToString("dd/MM/yyyy"); ;
                    ws.Cells[sat, 10].Value = "Sevk Bekleniyor";
                    int boxnoint = Int32.Parse(boxno);
                    int boxnoyeni = boxnoint + 1;
                    ws.Cells[sonrow, 1].Value = isemrit;
                    ws.Cells[sonrow, 2].Value = ucisemrit;
                    ws.Cells[sonrow, 3].Value = musteri;
                    ws.Cells[sonrow, 4].Value = DateTime.Now.ToString("dd/MM/yyyy");
                    ws.Cells[sonrow, 5].Value = termin;
                    ws.Cells[sonrow, 6].Value = projeadi;
                    ws.Cells[sonrow, 7].Value = boxnoyeni.ToString();
                    ws.Cells[sonrow, 10].Value = "Malzemeler hazırlanıyor";
                    ws.Cells[sonrow, 12].Value = notlar;
                    Outlook.Application oApp = new Outlook.Application();
                    string[] x1 = isemrit.Split('/');
                    string isemrid = string.Join("-", x1);
                    string yol = ("Z:\\TM-KLT TAKIP\\Malzeme listeleri\\" + isemrid + " " + musteri + "-" + projeadi + ".xlsx");
                    Outlook.MailItem oMsg = (Outlook.MailItem)oApp.CreateItem(Outlook.OlItemType.olMailItem);
                    oMsg.HTMLBody = "İş emri malzeme listesi:  " + " <a href= " + '\u0022' + yol + '\u0022' + ">" + isemrit + "</a>" + "<br>" + "<br>" + "Lütfen gönderilecek malzeme listesini yalnızca yukarıdaki linkten güncelleyiniz." + "<br>" + "Bu iş emrine diğer sandık bilgileri de eklenecektir." + "<br>" + "<br>" +
                                    "<b>İş emri numarası: </b>" + isemrit + "<br>" +
                                    "<b>Sandık boyutları:    </b>" + boxbilgisis + "<br>" +
                                    "<b>Notlar: </b>" + ws.Cells[sat, 12].Value.ToString();
                    //Subject line
                    oMsg.Subject = isemrit + " " + musteri +  " İş emrine ait sandık hazırlanmıştır";
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
                    if (ws.Cells[sonrow, 1].Value != null) break;
                }
            }
            ws.Cells["A1:L1"].AutoFilter = true;
            ws.Cells.AutoFitColumns(0);  //Autofit columns for all cells
            pck.Save();
            // Create the Outlook application.
            
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
            tbl.DefaultView.RowFilter = "Durum = '" + durum + "'";
            metroGrid1.DataSource = tbl;
            return;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            boxbilgisiekle();
            comboBox1.Items.Clear();
            acikisemirleri();
            comboBox1.Items.Clear();
            comboBox1.Text = "";
            foreach (DataGridViewRow row in metroGrid1.Rows)
            {
                comboBox1.Items.Add(row.Cells[0].Value.ToString());
            }
            this.Close();
        }

        private void Form2_FormClosing(object sender, FormClosingEventArgs e)
        {
            //Anasayfa.dene();
        }

    }
}