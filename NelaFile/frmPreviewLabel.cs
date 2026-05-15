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
    public partial class frmPreviewLabel : Form
    {
        public string PictureFolderFile { get; set; }
        public frmPreviewLabel()
        {
            InitializeComponent();
        }
        private void frmPreviewLabel_Load(object sender, EventArgs e)
        {
            Bitmap btmp = new Bitmap(PictureFolderFile);


            pictureBoxLabelPreview.Image = (Image)btmp; 


        }
    }
}
