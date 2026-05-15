namespace NelaFile
{
    partial class frmElektraData
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmElektraData));
            this.panelBottom = new System.Windows.Forms.Panel();
            this.buttonCancel = new System.Windows.Forms.Button();
            this.panel1 = new System.Windows.Forms.Panel();
            this.dataGridViewElektra = new System.Windows.Forms.DataGridView();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.textBox_C11_PACKAGE_NUM1 = new System.Windows.Forms.TextBox();
            this.textBox_C13_PACKAGE_NUM2 = new System.Windows.Forms.TextBox();
            this.button_CreateNewOrder = new System.Windows.Forms.Button();
            this.panelBottom.SuspendLayout();
            this.panel1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridViewElektra)).BeginInit();
            this.groupBox1.SuspendLayout();
            this.SuspendLayout();
            // 
            // panelBottom
            // 
            this.panelBottom.Controls.Add(this.buttonCancel);
            this.panelBottom.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.panelBottom.Location = new System.Drawing.Point(0, 414);
            this.panelBottom.Name = "panelBottom";
            this.panelBottom.Size = new System.Drawing.Size(1325, 48);
            this.panelBottom.TabIndex = 2;
            // 
            // buttonCancel
            // 
            this.buttonCancel.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.buttonCancel.Image = ((System.Drawing.Image)(resources.GetObject("buttonCancel.Image")));
            this.buttonCancel.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.buttonCancel.Location = new System.Drawing.Point(1238, 13);
            this.buttonCancel.Name = "buttonCancel";
            this.buttonCancel.Size = new System.Drawing.Size(75, 23);
            this.buttonCancel.TabIndex = 0;
            this.buttonCancel.Text = "Zavřít";
            this.buttonCancel.UseVisualStyleBackColor = true;
            this.buttonCancel.Click += new System.EventHandler(this.buttonCancel_Click);
            // 
            // panel1
            // 
            this.panel1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.panel1.Controls.Add(this.groupBox1);
            this.panel1.Controls.Add(this.dataGridViewElektra);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel1.ForeColor = System.Drawing.SystemColors.ScrollBar;
            this.panel1.Location = new System.Drawing.Point(0, 0);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(1325, 414);
            this.panel1.TabIndex = 3;
            // 
            // dataGridViewElektra
            // 
            this.dataGridViewElektra.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridViewElektra.Dock = System.Windows.Forms.DockStyle.Top;
            this.dataGridViewElektra.Location = new System.Drawing.Point(0, 0);
            this.dataGridViewElektra.Name = "dataGridViewElektra";
            this.dataGridViewElektra.Size = new System.Drawing.Size(1323, 189);
            this.dataGridViewElektra.TabIndex = 0;
            this.dataGridViewElektra.SelectionChanged += new System.EventHandler(this.dataGridViewElektra_SelectionChanged);
            // 
            // groupBox1
            // 
            this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left)));
            this.groupBox1.Controls.Add(this.button_CreateNewOrder);
            this.groupBox1.Controls.Add(this.textBox_C13_PACKAGE_NUM2);
            this.groupBox1.Controls.Add(this.textBox_C11_PACKAGE_NUM1);
            this.groupBox1.Controls.Add(this.label2);
            this.groupBox1.Controls.Add(this.label1);
            this.groupBox1.Location = new System.Drawing.Point(3, 195);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(276, 211);
            this.groupBox1.TabIndex = 1;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Znovu vyrobit";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(238)));
            this.label1.ForeColor = System.Drawing.Color.Red;
            this.label1.Location = new System.Drawing.Point(7, 26);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(143, 13);
            this.label1.TabIndex = 0;
            this.label1.Text = "Ood kartonu č. (včetně)";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(238)));
            this.label2.ForeColor = System.Drawing.Color.Red;
            this.label2.Location = new System.Drawing.Point(7, 52);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(128, 13);
            this.label2.TabIndex = 1;
            this.label2.Text = "Po karton č. (včetně)";
            // 
            // textBox_C11_PACKAGE_NUM1
            // 
            this.textBox_C11_PACKAGE_NUM1.Location = new System.Drawing.Point(159, 19);
            this.textBox_C11_PACKAGE_NUM1.Name = "textBox_C11_PACKAGE_NUM1";
            this.textBox_C11_PACKAGE_NUM1.Size = new System.Drawing.Size(86, 20);
            this.textBox_C11_PACKAGE_NUM1.TabIndex = 2;
            // 
            // textBox_C13_PACKAGE_NUM2
            // 
            this.textBox_C13_PACKAGE_NUM2.Location = new System.Drawing.Point(159, 45);
            this.textBox_C13_PACKAGE_NUM2.Name = "textBox_C13_PACKAGE_NUM2";
            this.textBox_C13_PACKAGE_NUM2.Size = new System.Drawing.Size(86, 20);
            this.textBox_C13_PACKAGE_NUM2.TabIndex = 3;
            // 
            // button_CreateNewOrder
            // 
            this.button_CreateNewOrder.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.button_CreateNewOrder.ForeColor = System.Drawing.SystemColors.WindowText;
            this.button_CreateNewOrder.Image = ((System.Drawing.Image)(resources.GetObject("button_CreateNewOrder.Image")));
            this.button_CreateNewOrder.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.button_CreateNewOrder.Location = new System.Drawing.Point(71, 143);
            this.button_CreateNewOrder.Name = "button_CreateNewOrder";
            this.button_CreateNewOrder.Size = new System.Drawing.Size(111, 37);
            this.button_CreateNewOrder.TabIndex = 4;
            this.button_CreateNewOrder.Text = "Vytvořit VO";
            this.button_CreateNewOrder.UseVisualStyleBackColor = true;
            this.button_CreateNewOrder.Click += new System.EventHandler(this.button_CreateNewOrder_Click);
            // 
            // frmElektraData
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1325, 462);
            this.Controls.Add(this.panel1);
            this.Controls.Add(this.panelBottom);
            this.Name = "frmElektraData";
            this.Text = "Elektra A1135";
            this.Load += new System.EventHandler(this.frmElektraData_Load);
            this.panelBottom.ResumeLayout(false);
            this.panel1.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridViewElektra)).EndInit();
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Panel panelBottom;
        private System.Windows.Forms.Button buttonCancel;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.DataGridView dataGridViewElektra;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.TextBox textBox_C13_PACKAGE_NUM2;
        private System.Windows.Forms.TextBox textBox_C11_PACKAGE_NUM1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button button_CreateNewOrder;
    }
}