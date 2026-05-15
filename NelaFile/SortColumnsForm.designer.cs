namespace NelaFile
{
    partial class SortColumnsForm
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(SortColumnsForm));
            this.btnAddLevel = new System.Windows.Forms.Button();
            this.btnDeleteLevel = new System.Windows.Forms.Button();
            this.btnClose = new System.Windows.Forms.Button();
            this.dgvSortColumns = new System.Windows.Forms.DataGridView();
            this.btnSort = new System.Windows.Forms.Button();
            this.chkCloseSort = new System.Windows.Forms.CheckBox();
            this.btnReset = new System.Windows.Forms.Button();
            this.cmbField = new System.Windows.Forms.DataGridViewComboBoxColumn();
            this.cmbOrder = new System.Windows.Forms.DataGridViewComboBoxColumn();
            ((System.ComponentModel.ISupportInitialize)(this.dgvSortColumns)).BeginInit();
            this.SuspendLayout();
            // 
            // btnAddLevel
            // 
            this.btnAddLevel.Image = ((System.Drawing.Image)(resources.GetObject("btnAddLevel.Image")));
            this.btnAddLevel.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btnAddLevel.Location = new System.Drawing.Point(12, 12);
            this.btnAddLevel.Name = "btnAddLevel";
            this.btnAddLevel.Size = new System.Drawing.Size(82, 23);
            this.btnAddLevel.TabIndex = 0;
            this.btnAddLevel.Text = "Add Level";
            this.btnAddLevel.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.btnAddLevel.UseVisualStyleBackColor = true;
            this.btnAddLevel.Click += new System.EventHandler(this.btnAddLevel_Click);
            // 
            // btnDeleteLevel
            // 
            this.btnDeleteLevel.Image = ((System.Drawing.Image)(resources.GetObject("btnDeleteLevel.Image")));
            this.btnDeleteLevel.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btnDeleteLevel.Location = new System.Drawing.Point(100, 12);
            this.btnDeleteLevel.Name = "btnDeleteLevel";
            this.btnDeleteLevel.Size = new System.Drawing.Size(94, 23);
            this.btnDeleteLevel.TabIndex = 1;
            this.btnDeleteLevel.Text = "Delete Level";
            this.btnDeleteLevel.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.btnDeleteLevel.UseVisualStyleBackColor = true;
            this.btnDeleteLevel.Click += new System.EventHandler(this.btnDeleteLevel_Click);
            // 
            // btnClose
            // 
            this.btnClose.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btnClose.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.btnClose.Location = new System.Drawing.Point(249, 227);
            this.btnClose.Name = "btnClose";
            this.btnClose.Size = new System.Drawing.Size(75, 23);
            this.btnClose.TabIndex = 3;
            this.btnClose.Text = "Close";
            this.btnClose.UseVisualStyleBackColor = true;
            // 
            // dgvSortColumns
            // 
            this.dgvSortColumns.AllowUserToAddRows = false;
            this.dgvSortColumns.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)
                        | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.dgvSortColumns.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgvSortColumns.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.cmbField,
            this.cmbOrder});
            this.dgvSortColumns.Location = new System.Drawing.Point(12, 41);
            this.dgvSortColumns.Name = "dgvSortColumns";
            this.dgvSortColumns.RowHeadersWidth = 90;
            this.dgvSortColumns.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing;
            this.dgvSortColumns.Size = new System.Drawing.Size(312, 180);
            this.dgvSortColumns.TabIndex = 4;
            this.dgvSortColumns.CellClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridViewSortColumns_CellClick);
            // 
            // btnSort
            // 
            this.btnSort.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btnSort.Location = new System.Drawing.Point(168, 227);
            this.btnSort.Name = "btnSort";
            this.btnSort.Size = new System.Drawing.Size(75, 23);
            this.btnSort.TabIndex = 5;
            this.btnSort.Text = "Sort";
            this.btnSort.UseVisualStyleBackColor = true;
            this.btnSort.Click += new System.EventHandler(this.btnSort_Click);
            // 
            // chkCloseSort
            // 
            this.chkCloseSort.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.chkCloseSort.AutoSize = true;
            this.chkCloseSort.Checked = true;
            this.chkCloseSort.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkCloseSort.Location = new System.Drawing.Point(12, 233);
            this.chkCloseSort.Name = "chkCloseSort";
            this.chkCloseSort.Size = new System.Drawing.Size(89, 17);
            this.chkCloseSort.TabIndex = 7;
            this.chkCloseSort.Text = "Close on Sort";
            this.chkCloseSort.UseVisualStyleBackColor = true;
            // 
            // btnReset
            // 
            this.btnReset.Image = ((System.Drawing.Image)(resources.GetObject("btnReset.Image")));
            this.btnReset.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btnReset.Location = new System.Drawing.Point(200, 12);
            this.btnReset.Name = "btnReset";
            this.btnReset.Size = new System.Drawing.Size(87, 23);
            this.btnReset.TabIndex = 8;
            this.btnReset.Text = "Reset";
            this.btnReset.UseVisualStyleBackColor = true;
            this.btnReset.Click += new System.EventHandler(this.btnReset_Click);
            // 
            // cmbField
            // 
            this.cmbField.FillWeight = 139.0863F;
            this.cmbField.HeaderText = "Column";
            this.cmbField.Name = "cmbField";
            this.cmbField.Resizable = System.Windows.Forms.DataGridViewTriState.True;
            // 
            // cmbOrder
            // 
            this.cmbOrder.FillWeight = 60.9137F;
            this.cmbOrder.HeaderText = "Order";
            this.cmbOrder.Items.AddRange(new object[] {
            "Ascending",
            "Descending"});
            this.cmbOrder.Name = "cmbOrder";
            this.cmbOrder.Resizable = System.Windows.Forms.DataGridViewTriState.False;
            this.cmbOrder.Width = 97;
            // 
            // SortColumnsForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.CancelButton = this.btnClose;
            this.ClientSize = new System.Drawing.Size(336, 262);
            this.Controls.Add(this.btnReset);
            this.Controls.Add(this.chkCloseSort);
            this.Controls.Add(this.btnSort);
            this.Controls.Add(this.dgvSortColumns);
            this.Controls.Add(this.btnClose);
            this.Controls.Add(this.btnDeleteLevel);
            this.Controls.Add(this.btnAddLevel);
            this.MaximizeBox = false;
            this.MaximumSize = new System.Drawing.Size(800, 500);
            this.MinimizeBox = false;
            this.MinimumSize = new System.Drawing.Size(352, 300);
            this.Name = "SortColumnsForm";
            this.ShowInTaskbar = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "Sort";
            this.Load += new System.EventHandler(this.SortColumnsForm_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dgvSortColumns)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnAddLevel;
        private System.Windows.Forms.Button btnDeleteLevel;
        private System.Windows.Forms.Button btnClose;
        private System.Windows.Forms.DataGridView dgvSortColumns;
        private System.Windows.Forms.Button btnSort;
        private System.Windows.Forms.CheckBox chkCloseSort;
        private System.Windows.Forms.Button btnReset;
        private System.Windows.Forms.DataGridViewComboBoxColumn cmbField;
        private System.Windows.Forms.DataGridViewComboBoxColumn cmbOrder;
    }
}