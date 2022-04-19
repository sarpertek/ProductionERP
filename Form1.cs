using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.SQLite;
using MetroFramework.Forms;
using MetroFramework;
using System.Reflection;

namespace _2015
{
    public partial class Form1 : MetroForm
    {
        public Form1()
        {
            InitializeComponent();
            this.TopMost = true;
            this.BringToFront();
            this.Focus();
            Init_Data();
        }
        private void Form1_Load(object sender, EventArgs e)
        {
            this.ActiveControl = metroTextBox1;
        }

        private void metroButton1_Click(object sender, EventArgs e)
        {
            SQLiteConnection baglan = new SQLiteConnection(baglanti.database);
            baglan.Open();
            string sql = "SELECT departman FROM kullanicilar WHERE kullaniciadi=@kadi and sifre=@sifre";
            SQLiteParameter prm1 = new SQLiteParameter("kadi", metroTextBox1.Text);
            SQLiteParameter prm2 = new SQLiteParameter("sifre", metroTextBox2.Text);
            SQLiteCommand cmd = new SQLiteCommand(sql, baglan);
            cmd.Parameters.Add(prm1);
            cmd.Parameters.Add(prm2);
            DataTable dt = new DataTable();
            SQLiteDataAdapter da = new SQLiteDataAdapter(cmd);
            string departman = (string)cmd.ExecuteScalar();
            da.Fill(dt);           
            baglan.Close();

            if (dt.Rows.Count > 0)
            {
                Main main = new Main();
                Anasayfa.Instance.metroTile1.Enabled = true;
                Anasayfa.Instance.metroTile3.Enabled = true;
                Anasayfa.Instance.metroTile4.Enabled = true;
                Anasayfa.Instance.metroTile5.Enabled = true;
                if (departman == "Dış Ticaret")
                {
                    Anasayfa.Instance.metroTile6.Enabled = true;
                }
                else
                {
                    Anasayfa.Instance.metroTile6.Enabled = false;
                }
                anasayfa2.Instance.metroTile2.Enabled = true;
                Save_Data();
                main.Show();
                this.Hide();
            }
            else
            {
                MessageBox.Show("Hatalı bir giriş yaptınız lütfen tekrar deneyiniz.");
            }

        }

        private void devambutton_Click(object sender, EventArgs e)
        {
            DialogResult dialogResult = MessageBox.Show("Giriş yapmadan devam ederseniz düzenleme yapamazsınız. Giriş yapmadan devam etmek istiyor musunuz?", "Uyarı", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (dialogResult == DialogResult.Yes)
            {
                Main main = new Main();
                Anasayfa.Instance.metroTile1.Enabled = false;
                Anasayfa.Instance.metroTile3.Enabled = false;
                Anasayfa.Instance.metroTile4.Enabled = false;
                Anasayfa.Instance.metroTile5.Enabled = false;
                Anasayfa.Instance.metroTile6.Enabled = false;
                anasayfa2.Instance.metroTile2.Enabled = false;
                main.metroLabel1.Text = "Düzenleme yapabilmek için giriş yapmalısınız. Giriş yapmak için tıklayın.";
                main.Show();
                this.Hide();
            }
            else if (dialogResult == DialogResult.No)
            {
            }
        }
        private void Init_Data()
        {
            if (Properties.Settings.Default.UserName != string.Empty)
            {
                if (Properties.Settings.Default.Remme == "yes")
                {
                    metroTextBox1.Text = Properties.Settings.Default.UserName;
                    metroTextBox2.Text = Properties.Settings.Default.Password;
                    checkRememer.Checked = true;
                }
                else
                {
                    metroTextBox1.Text = "";
                }
            }
        }
        private void Save_Data()
        {
            if (checkRememer.Checked)
            {
                Properties.Settings.Default.UserName = metroTextBox1.Text;
                Properties.Settings.Default.Password = metroTextBox2.Text;
                Properties.Settings.Default.Remme = "yes";
                Properties.Settings.Default.Save();
            }
            else
            {
                Properties.Settings.Default.UserName = "";
                Properties.Settings.Default.Password = "";
                Properties.Settings.Default.Remme = "no";
                Properties.Settings.Default.Save();
            }
        }

        private void metroTextBox2_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                girisbutton.PerformClick();
            }
        }

        private void metroTextBox1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                metroTextBox2.Focus();
            }
        }
    }
}
