using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace NelaFile
{
    public partial class frmFileView : Form
    {
        public string FileName { get; set; }
#pragma warning disable CS0169 // The field 'frmFileView._Words' is never used
        private List<string> _Words;
#pragma warning restore CS0169 // The field 'frmFileView._Words' is never used
        private int _Pozice = 1;

        public frmFileView()
        {
            InitializeComponent();
        }

        private void frmFileView_Load(object sender, EventArgs e)
        {
            textBoxFile.Text = System.IO.File.ReadAllText(this.FileName,Encoding.Default);
        }

        private void toolStripTextBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar.ToString() == "\r")
            {
                string sPom = string.Empty;
                int pos = textBoxFile.Text.IndexOf(toolStripTextBox1.Text, _Pozice);
                //while (pos != -1)
                if (pos != -1)
                {
                    //textBoxFile.SelectionStart = pos;
                    textBoxFile.Select(pos, toolStripTextBox1.Text.Length);
                    _Pozice = pos+1;
                    //textBoxFile.SelectionLength = toolStripTextBox1.Text.Length;
                    Console.WriteLine(_Pozice.ToString());
                    //pos = textBoxFile.Text.IndexOf(toolStripTextBox1.Text,pos+1);
                }
                textBoxFile.Focus();
            }
            else
            {
                //_Pozice = 1;
            }

        }

        private void toolStripTextBox1_Enter(object sender, EventArgs e)
        {
        }

        private void toolStripTextBox1_Click(object sender, EventArgs e)
        {
            _Pozice = 0;
        }

    }
}
