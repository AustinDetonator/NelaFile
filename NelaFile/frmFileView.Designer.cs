namespace NelaFile
{
    partial class frmFileView
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
            this.textBoxFile = new System.Windows.Forms.TextBox();
            this.contextMenuStrip1 = new System.Windows.Forms.ContextMenuStrip(this.components);
            this.ToolStripMenuItem_Hledej = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripTextBox1 = new System.Windows.Forms.ToolStripTextBox();
            this.contextMenuStrip1.SuspendLayout();
            this.SuspendLayout();
            // 
            // textBoxFile
            // 
            this.textBoxFile.ContextMenuStrip = this.contextMenuStrip1;
            this.textBoxFile.Dock = System.Windows.Forms.DockStyle.Fill;
            this.textBoxFile.Font = new System.Drawing.Font("Courier New", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(238)));
            this.textBoxFile.Location = new System.Drawing.Point(0, 0);
            this.textBoxFile.Multiline = true;
            this.textBoxFile.Name = "textBoxFile";
            this.textBoxFile.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.textBoxFile.Size = new System.Drawing.Size(651, 722);
            this.textBoxFile.TabIndex = 0;
            this.textBoxFile.TabStop = false;
            // 
            // contextMenuStrip1
            // 
            this.contextMenuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.ToolStripMenuItem_Hledej});
            this.contextMenuStrip1.Name = "contextMenuStrip1";
            this.contextMenuStrip1.Size = new System.Drawing.Size(153, 48);
            // 
            // ToolStripMenuItem_Hledej
            // 
            this.ToolStripMenuItem_Hledej.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.toolStripTextBox1});
            this.ToolStripMenuItem_Hledej.Name = "ToolStripMenuItem_Hledej";
            this.ToolStripMenuItem_Hledej.Size = new System.Drawing.Size(152, 22);
            this.ToolStripMenuItem_Hledej.Text = "Hledej";
            // 
            // toolStripTextBox1
            // 
            this.toolStripTextBox1.Name = "toolStripTextBox1";
            this.toolStripTextBox1.Size = new System.Drawing.Size(100, 23);
            this.toolStripTextBox1.Enter += new System.EventHandler(this.toolStripTextBox1_Enter);
            this.toolStripTextBox1.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.toolStripTextBox1_KeyPress);
            this.toolStripTextBox1.Click += new System.EventHandler(this.toolStripTextBox1_Click);
            // 
            // frmFileView
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(651, 722);
            this.Controls.Add(this.textBoxFile);
            this.Name = "frmFileView";
            this.Text = "frmFileView";
            this.Load += new System.EventHandler(this.frmFileView_Load);
            this.contextMenuStrip1.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox textBoxFile;
        private System.Windows.Forms.ContextMenuStrip contextMenuStrip1;
        private System.Windows.Forms.ToolStripMenuItem ToolStripMenuItem_Hledej;
        private System.Windows.Forms.ToolStripTextBox toolStripTextBox1;
    }
}