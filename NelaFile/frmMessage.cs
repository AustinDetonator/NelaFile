using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace NelaFile
{
    public partial class frmMessage : Form
    {
        public string Message { get; set; }
        //
        public frmMessage()
        {
            InitializeComponent();
        }

        private void buttonOk_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void frmMessage_Load(object sender, EventArgs e)
        {
            textBoxMessage.Text = this.Message;
        }
    }
}
