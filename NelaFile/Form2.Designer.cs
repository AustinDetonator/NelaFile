namespace NelaFile
{
    partial class Form2
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form2));
            this.workbookView1 = new SpreadsheetGear.Windows.Forms.WorkbookView();
            this.SuspendLayout();
            // 
            // workbookView1
            // 
            this.workbookView1.FormulaBar = null;
            this.workbookView1.Location = new System.Drawing.Point(7, 4);
            this.workbookView1.Name = "workbookView1";
            this.workbookView1.Size = new System.Drawing.Size(742, 414);
            this.workbookView1.TabIndex = 0;
            this.workbookView1.WorkbookSetState = resources.GetString("workbookView1.WorkbookSetState");
            // 
            // Form2
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(761, 421);
            this.Controls.Add(this.workbookView1);
            this.Name = "Form2";
            this.Text = "Form2";
            this.ResumeLayout(false);

        }

        #endregion

        private SpreadsheetGear.Windows.Forms.WorkbookView workbookView1;
    }
}