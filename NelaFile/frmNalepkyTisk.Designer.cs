namespace NelaFile
{
    partial class frmNalepkyTisk
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
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmNalepkyTisk));
            this.panelMain = new System.Windows.Forms.Panel();
            this.buttonLabelsCreate = new System.Windows.Forms.Button();
            this.groupBoxRadky = new System.Windows.Forms.GroupBox();
            this.dataGridViewLabels = new System.Windows.Forms.DataGridView();
            this.contextMenuStripLabels = new System.Windows.Forms.ContextMenuStrip(this.components);
            this.toolStripMenuItemLabels_Preview = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripMenuItemLabels_Print = new System.Windows.Forms.ToolStripMenuItem();
            this.panelBottom = new System.Windows.Forms.Panel();
            this.buttonCancel = new System.Windows.Forms.Button();
            this.panelTop = new System.Windows.Forms.Panel();
            this.buttonPrintLabel = new System.Windows.Forms.Button();
            this.labelCustomer = new System.Windows.Forms.Label();
            this.labelTiskárna = new System.Windows.Forms.Label();
            this.buttonLoty = new System.Windows.Forms.Button();
            this.btnLabelPreview = new System.Windows.Forms.Button();
            this.comboBoxPrinters = new System.Windows.Forms.ComboBox();
            this.textBoxInfoZO = new System.Windows.Forms.TextBox();
            this.labelInfoZO = new System.Windows.Forms.Label();
            this.groupBoxNalepky = new System.Windows.Forms.GroupBox();
            this.dataGridViewBTW = new System.Windows.Forms.DataGridView();
            this.toolTip1 = new System.Windows.Forms.ToolTip(this.components);
            this.imageList1 = new System.Windows.Forms.ImageList(this.components);
            this.bindingSource1 = new System.Windows.Forms.BindingSource(this.components);
            this.panelMain.SuspendLayout();
            this.groupBoxRadky.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridViewLabels)).BeginInit();
            this.contextMenuStripLabels.SuspendLayout();
            this.panelBottom.SuspendLayout();
            this.panelTop.SuspendLayout();
            this.groupBoxNalepky.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridViewBTW)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.bindingSource1)).BeginInit();
            this.SuspendLayout();
            // 
            // panelMain
            // 
            this.panelMain.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.panelMain.Controls.Add(this.buttonLabelsCreate);
            this.panelMain.Controls.Add(this.groupBoxRadky);
            this.panelMain.Location = new System.Drawing.Point(0, 61);
            this.panelMain.Name = "panelMain";
            this.panelMain.Size = new System.Drawing.Size(1546, 676);
            this.panelMain.TabIndex = 0;
            // 
            // buttonLabelsCreate
            // 
            this.buttonLabelsCreate.Image = ((System.Drawing.Image)(resources.GetObject("buttonLabelsCreate.Image")));
            this.buttonLabelsCreate.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.buttonLabelsCreate.Location = new System.Drawing.Point(12, 400);
            this.buttonLabelsCreate.Name = "buttonLabelsCreate";
            this.buttonLabelsCreate.Size = new System.Drawing.Size(114, 23);
            this.buttonLabelsCreate.TabIndex = 4;
            this.buttonLabelsCreate.Text = "Nálepky";
            this.buttonLabelsCreate.UseVisualStyleBackColor = true;
            this.buttonLabelsCreate.Click += new System.EventHandler(this.buttonLabelsCreate_Click);
            // 
            // groupBoxRadky
            // 
            this.groupBoxRadky.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.groupBoxRadky.Controls.Add(this.dataGridViewLabels);
            this.groupBoxRadky.Location = new System.Drawing.Point(256, 3);
            this.groupBoxRadky.Name = "groupBoxRadky";
            this.groupBoxRadky.Size = new System.Drawing.Size(1287, 670);
            this.groupBoxRadky.TabIndex = 1;
            this.groupBoxRadky.TabStop = false;
            this.groupBoxRadky.Text = "Seznam nálepek";
            // 
            // dataGridViewLabels
            // 
            this.dataGridViewLabels.AllowUserToAddRows = false;
            this.dataGridViewLabels.AllowUserToDeleteRows = false;
            this.dataGridViewLabels.AllowUserToResizeRows = false;
            this.dataGridViewLabels.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells;
            this.dataGridViewLabels.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridViewLabels.ContextMenuStrip = this.contextMenuStripLabels;
            this.dataGridViewLabels.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dataGridViewLabels.Location = new System.Drawing.Point(3, 16);
            this.dataGridViewLabels.Name = "dataGridViewLabels";
            this.dataGridViewLabels.Size = new System.Drawing.Size(1281, 651);
            this.dataGridViewLabels.TabIndex = 0;
            this.dataGridViewLabels.DoubleClick += new System.EventHandler(this.dataGridViewLabels_DoubleClick);
            // 
            // contextMenuStripLabels
            // 
            this.contextMenuStripLabels.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.toolStripMenuItemLabels_Preview,
            this.toolStripMenuItemLabels_Print});
            this.contextMenuStripLabels.Name = "contextMenuStripLabels";
            this.contextMenuStripLabels.Size = new System.Drawing.Size(197, 48);
            // 
            // toolStripMenuItemLabels_Preview
            // 
            this.toolStripMenuItemLabels_Preview.Name = "toolStripMenuItemLabels_Preview";
            this.toolStripMenuItemLabels_Preview.Size = new System.Drawing.Size(196, 22);
            this.toolStripMenuItemLabels_Preview.Text = "Náhled";
            this.toolStripMenuItemLabels_Preview.Click += new System.EventHandler(this.toolStripMenuItemLabels_Preview_Click);
            // 
            // toolStripMenuItemLabels_Print
            // 
            this.toolStripMenuItemLabels_Print.Name = "toolStripMenuItemLabels_Print";
            this.toolStripMenuItemLabels_Print.Size = new System.Drawing.Size(196, 22);
            this.toolStripMenuItemLabels_Print.Text = "Tisk vybraných nálepek";
            this.toolStripMenuItemLabels_Print.Click += new System.EventHandler(this.toolStripMenuItemLabels_Print_Click);
            // 
            // panelBottom
            // 
            this.panelBottom.Controls.Add(this.buttonCancel);
            this.panelBottom.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.panelBottom.Location = new System.Drawing.Point(0, 743);
            this.panelBottom.Name = "panelBottom";
            this.panelBottom.Size = new System.Drawing.Size(1546, 48);
            this.panelBottom.TabIndex = 1;
            // 
            // buttonCancel
            // 
            this.buttonCancel.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.buttonCancel.Image = ((System.Drawing.Image)(resources.GetObject("buttonCancel.Image")));
            this.buttonCancel.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.buttonCancel.Location = new System.Drawing.Point(1459, 13);
            this.buttonCancel.Name = "buttonCancel";
            this.buttonCancel.Size = new System.Drawing.Size(75, 23);
            this.buttonCancel.TabIndex = 0;
            this.buttonCancel.Text = "Zavřít";
            this.buttonCancel.UseVisualStyleBackColor = true;
            this.buttonCancel.Click += new System.EventHandler(this.buttonCancel_Click);
            // 
            // panelTop
            // 
            this.panelTop.Controls.Add(this.buttonPrintLabel);
            this.panelTop.Controls.Add(this.labelCustomer);
            this.panelTop.Controls.Add(this.labelTiskárna);
            this.panelTop.Controls.Add(this.buttonLoty);
            this.panelTop.Controls.Add(this.btnLabelPreview);
            this.panelTop.Controls.Add(this.comboBoxPrinters);
            this.panelTop.Controls.Add(this.textBoxInfoZO);
            this.panelTop.Controls.Add(this.labelInfoZO);
            this.panelTop.Dock = System.Windows.Forms.DockStyle.Top;
            this.panelTop.Location = new System.Drawing.Point(0, 0);
            this.panelTop.Name = "panelTop";
            this.panelTop.Size = new System.Drawing.Size(1546, 55);
            this.panelTop.TabIndex = 2;
            // 
            // buttonPrintLabel
            // 
            this.buttonPrintLabel.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.buttonPrintLabel.Image = ((System.Drawing.Image)(resources.GetObject("buttonPrintLabel.Image")));
            this.buttonPrintLabel.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.buttonPrintLabel.Location = new System.Drawing.Point(719, 16);
            this.buttonPrintLabel.Name = "buttonPrintLabel";
            this.buttonPrintLabel.Size = new System.Drawing.Size(95, 25);
            this.buttonPrintLabel.TabIndex = 10;
            this.buttonPrintLabel.Text = "Tisk";
            this.buttonPrintLabel.UseVisualStyleBackColor = true;
            this.buttonPrintLabel.Click += new System.EventHandler(this.buttonPrintLabel_Click);
            // 
            // labelCustomer
            // 
            this.labelCustomer.AutoSize = true;
            this.labelCustomer.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(238)));
            this.labelCustomer.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(0)))), ((int)(((byte)(192)))));
            this.labelCustomer.Location = new System.Drawing.Point(12, 29);
            this.labelCustomer.Name = "labelCustomer";
            this.labelCustomer.Size = new System.Drawing.Size(46, 17);
            this.labelCustomer.TabIndex = 9;
            this.labelCustomer.Text = "label1";
            // 
            // labelTiskárna
            // 
            this.labelTiskárna.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.labelTiskárna.AutoSize = true;
            this.labelTiskárna.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(238)));
            this.labelTiskárna.Location = new System.Drawing.Point(972, 24);
            this.labelTiskárna.Name = "labelTiskárna";
            this.labelTiskárna.Size = new System.Drawing.Size(67, 17);
            this.labelTiskárna.TabIndex = 8;
            this.labelTiskárna.Text = "Tiskárna:";
            // 
            // buttonLoty
            // 
            this.buttonLoty.Image = ((System.Drawing.Image)(resources.GetObject("buttonLoty.Image")));
            this.buttonLoty.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.buttonLoty.Location = new System.Drawing.Point(360, 16);
            this.buttonLoty.Name = "buttonLoty";
            this.buttonLoty.Size = new System.Drawing.Size(95, 25);
            this.buttonLoty.TabIndex = 7;
            this.buttonLoty.Text = "Loty";
            this.buttonLoty.UseVisualStyleBackColor = true;
            // 
            // btnLabelPreview
            // 
            this.btnLabelPreview.Image = ((System.Drawing.Image)(resources.GetObject("btnLabelPreview.Image")));
            this.btnLabelPreview.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btnLabelPreview.Location = new System.Drawing.Point(604, 16);
            this.btnLabelPreview.Name = "btnLabelPreview";
            this.btnLabelPreview.Size = new System.Drawing.Size(95, 25);
            this.btnLabelPreview.TabIndex = 6;
            this.btnLabelPreview.Text = "Náhled";
            this.btnLabelPreview.UseVisualStyleBackColor = true;
            this.btnLabelPreview.Click += new System.EventHandler(this.btnLabelPreview_Click);
            // 
            // comboBoxPrinters
            // 
            this.comboBoxPrinters.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.comboBoxPrinters.FormattingEnabled = true;
            this.comboBoxPrinters.Location = new System.Drawing.Point(1045, 20);
            this.comboBoxPrinters.Name = "comboBoxPrinters";
            this.comboBoxPrinters.Size = new System.Drawing.Size(216, 21);
            this.comboBoxPrinters.TabIndex = 2;
            this.comboBoxPrinters.SelectedIndexChanged += new System.EventHandler(this.comboBoxPrinters_SelectedIndexChanged);
            // 
            // textBoxInfoZO
            // 
            this.textBoxInfoZO.Location = new System.Drawing.Point(1344, 20);
            this.textBoxInfoZO.Name = "textBoxInfoZO";
            this.textBoxInfoZO.Size = new System.Drawing.Size(133, 20);
            this.textBoxInfoZO.TabIndex = 1;
            this.textBoxInfoZO.Visible = false;
            // 
            // labelInfoZO
            // 
            this.labelInfoZO.AutoSize = true;
            this.labelInfoZO.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(238)));
            this.labelInfoZO.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(0)))), ((int)(((byte)(192)))));
            this.labelInfoZO.Location = new System.Drawing.Point(12, 9);
            this.labelInfoZO.Name = "labelInfoZO";
            this.labelInfoZO.Size = new System.Drawing.Size(46, 17);
            this.labelInfoZO.TabIndex = 0;
            this.labelInfoZO.Text = "label1";
            // 
            // groupBoxNalepky
            // 
            this.groupBoxNalepky.Controls.Add(this.dataGridViewBTW);
            this.groupBoxNalepky.Location = new System.Drawing.Point(3, 64);
            this.groupBoxNalepky.Name = "groupBoxNalepky";
            this.groupBoxNalepky.Size = new System.Drawing.Size(250, 391);
            this.groupBoxNalepky.TabIndex = 0;
            this.groupBoxNalepky.TabStop = false;
            this.groupBoxNalepky.Text = "BTW";
            // 
            // dataGridViewBTW
            // 
            this.dataGridViewBTW.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridViewBTW.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dataGridViewBTW.Location = new System.Drawing.Point(3, 16);
            this.dataGridViewBTW.Name = "dataGridViewBTW";
            this.dataGridViewBTW.Size = new System.Drawing.Size(244, 372);
            this.dataGridViewBTW.TabIndex = 1;
            this.dataGridViewBTW.SelectionChanged += new System.EventHandler(this.dataGridViewBTW_SelectionChanged);
            this.dataGridViewBTW.DoubleClick += new System.EventHandler(this.dataGridViewBTW_DoubleClick);
            // 
            // imageList1
            // 
            this.imageList1.ColorDepth = System.Windows.Forms.ColorDepth.Depth8Bit;
            this.imageList1.ImageSize = new System.Drawing.Size(16, 16);
            this.imageList1.TransparentColor = System.Drawing.Color.Transparent;
            // 
            // bindingSource1
            // 
            this.bindingSource1.CurrentChanged += new System.EventHandler(this.bindingSource1_CurrentChanged);
            // 
            // frmNalepkyTisk
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1546, 791);
            this.Controls.Add(this.panelTop);
            this.Controls.Add(this.groupBoxNalepky);
            this.Controls.Add(this.panelBottom);
            this.Controls.Add(this.panelMain);
            this.Name = "frmNalepkyTisk";
            this.Text = "frmNalepkyTisk";
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.frmNalepkyTisk_FormClosed);
            this.Load += new System.EventHandler(this.frmNalepkyTisk_Load);
            this.panelMain.ResumeLayout(false);
            this.groupBoxRadky.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridViewLabels)).EndInit();
            this.contextMenuStripLabels.ResumeLayout(false);
            this.panelBottom.ResumeLayout(false);
            this.panelTop.ResumeLayout(false);
            this.panelTop.PerformLayout();
            this.groupBoxNalepky.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridViewBTW)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.bindingSource1)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Panel panelMain;
        private System.Windows.Forms.Panel panelBottom;
        private System.Windows.Forms.Button buttonCancel;
        private System.Windows.Forms.Panel panelTop;
        private System.Windows.Forms.GroupBox groupBoxRadky;
        private System.Windows.Forms.DataGridView dataGridViewLabels;
        private System.Windows.Forms.GroupBox groupBoxNalepky;
        private System.Windows.Forms.DataGridView dataGridViewBTW;
        private System.Windows.Forms.Button buttonLabelsCreate;
        private System.Windows.Forms.Label labelInfoZO;
        private System.Windows.Forms.TextBox textBoxInfoZO;
        private System.Windows.Forms.ToolTip toolTip1;
        private System.Windows.Forms.ContextMenuStrip contextMenuStripLabels;
        private System.Windows.Forms.ToolStripMenuItem toolStripMenuItemLabels_Preview;
        private System.Windows.Forms.ToolStripMenuItem toolStripMenuItemLabels_Print;
        private System.Windows.Forms.ComboBox comboBoxPrinters;
        private System.Windows.Forms.ImageList imageList1;
        private System.Windows.Forms.Button buttonLoty;
        private System.Windows.Forms.Button btnLabelPreview;
        private System.Windows.Forms.Label labelTiskárna;
        private System.Windows.Forms.Label labelCustomer;
        private System.Windows.Forms.Button buttonPrintLabel;
        private System.Windows.Forms.BindingSource bindingSource1;
    }
}