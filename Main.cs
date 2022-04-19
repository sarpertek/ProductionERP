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

namespace _2015
{
    public partial class Main : MetroForm
    {
        public Main()
        {
            InitializeComponent();
            this.StyleManager = metroStyleManager1;
        }

        private void Main_Load(object sender, EventArgs e)
        {
            if (!metroPanel1.Controls.Contains(Anasayfa.Instance))
            {
                metroPanel1.Controls.Add(Anasayfa.Instance);
                Anasayfa.Instance.BringToFront();
                Anasayfa.Instance.Dock = DockStyle.Fill;
            }
            else
                Anasayfa.Instance.BringToFront();
            Anasayfa.Instance.Dock = DockStyle.Fill;
            //metroTile1.Enabled = false;
            //metroTile2.Enabled = true;
        }

        private void metroTile2_Click(object sender, EventArgs e)
        {
            if (!metroPanel1.Controls.Contains(anasayfa2.Instance))
            {
                metroPanel1.Controls.Add(anasayfa2.Instance);
                anasayfa2.Instance.BringToFront();
                anasayfa2.Instance.Dock = DockStyle.Fill;
            }
            else
                anasayfa2.Instance.BringToFront();
            anasayfa2.Instance.Dock = DockStyle.Fill;
            
            //metroTile2.Enabled = false;
            //metroTile1.Enabled = true;
        }

        private void metroTile1_Click(object sender, EventArgs e)
        {
            if (!metroPanel1.Controls.Contains(Anasayfa.Instance))
            {
                metroPanel1.Controls.Add(Anasayfa.Instance);
                Anasayfa.Instance.BringToFront();
            }
            else
                Anasayfa.Instance.BringToFront();
            Anasayfa.Instance.Dock = DockStyle.Fill;
            //metroTile1.Enabled = false;
            //metroTile2.Enabled = true;

        }

        private void Main_FormClosing(object sender, System.Windows.Forms.FormClosingEventArgs e)
        {
            Application.Exit();
        }

        private void metroLabel1_Click(object sender, EventArgs e)
        {
            if (metroLabel1.Text == "Düzenleme yapabilmek için giriş yapmalısınız. Giriş yapmak için tıklayın.")
            {
                Form1 form1 = new Form1();
                form1.Show();
                this.Hide();
            }
        }

        private void metroButton1_Click(object sender, EventArgs e)
        {
            if (!metroPanel1.Controls.Contains(Teklif.Instance))
            {
                metroPanel1.Controls.Add(Teklif.Instance);
                Teklif.Instance.BringToFront();
                Teklif.Instance.Dock = DockStyle.Fill;
            }
            else
                Teklif.Instance.BringToFront();
            Teklif.Instance.Dock = DockStyle.Fill;
            //metroTile1.Enabled = false;
            //metroTile2.Enabled = true;
        }

        private void metroButton2_Click(object sender, EventArgs e)
        {

        }

        private void metroButton2_Click_1(object sender, EventArgs e)
        {
            if (!metroPanel1.Controls.Contains(UC.Instance))
            {
                metroPanel1.Controls.Add(UC.Instance);
                UC.Instance.BringToFront();
                UC.Instance.Dock = DockStyle.Fill;
            }
            else
                UC.Instance.BringToFront();
                UC.Instance.Dock = DockStyle.Fill;
            //metroTile2.Enabled = false;
            //metroTile1.Enabled = true;
        }

        private void metroButton3_Click(object sender, EventArgs e)
        {
            metroStyleManager1.Theme = MetroFramework.MetroThemeStyle.Light;
        }

        private void metroButton4_Click(object sender, EventArgs e)
        {
            metroStyleManager1.Theme = MetroFramework.MetroThemeStyle.Dark;
            MessageBox.Show("anasayfa2");
        }
    }
}
