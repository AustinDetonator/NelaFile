namespace NelaFile
{
    partial class frmPreviewLabel
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmPreviewLabel));
            this.groupBoxLabelPreview = new System.Windows.Forms.GroupBox();
            this.pictureBoxLabelPreview = new System.Windows.Forms.PictureBox();
            this.groupBoxLabelPreview.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBoxLabelPreview)).BeginInit();
            this.SuspendLayout();
            // 
            // groupBoxLabelPreview
            // 
            this.groupBoxLabelPreview.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.groupBoxLabelPreview.Controls.Add(this.pictureBoxLabelPreview);
            this.groupBoxLabelPreview.Dock = System.Windows.Forms.DockStyle.Fill;
            this.groupBoxLabelPreview.Location = new System.Drawing.Point(0, 0);
            this.groupBoxLabelPreview.Name = "groupBoxLabelPreview";
            this.groupBoxLabelPreview.Size = new System.Drawing.Size(467, 671);
            this.groupBoxLabelPreview.TabIndex = 4;
            this.groupBoxLabelPreview.TabStop = false;
            this.groupBoxLabelPreview.Text = "Náhled";
            // 
            // pictureBoxLabelPreview
            // 
            this.pictureBoxLabelPreview.Dock = System.Windows.Forms.DockStyle.Fill;
            this.pictureBoxLabelPreview.Location = new System.Drawing.Point(3, 16);
            this.pictureBoxLabelPreview.Name = "pictureBoxLabelPreview";
            this.pictureBoxLabelPreview.Size = new System.Drawing.Size(461, 652);
            this.pictureBoxLabelPreview.TabIndex = 0;
            this.pictureBoxLabelPreview.TabStop = false;
            // 
            // frmPreviewLabel
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(467, 671);
            this.Controls.Add(this.groupBoxLabelPreview);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "frmPreviewLabel";
            this.Text = "Náhled nálepky";
            this.Load += new System.EventHandler(this.frmPreviewLabel_Load);
            this.groupBoxLabelPreview.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.pictureBoxLabelPreview)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.GroupBox groupBoxLabelPreview;
        private System.Windows.Forms.PictureBox pictureBoxLabelPreview;
    }
}