using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace _2015
{
    public partial class uretim : MetroFramework.Controls.MetroUserControl
    {
        private static uretim _instance;
        public static uretim Instance
        {
            get
            {
                if (_instance == null)
                    _instance = new uretim();
                return _instance;
            }
        }
        public uretim()
        {
            InitializeComponent();
        }
    }
}
