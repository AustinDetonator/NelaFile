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
    public partial class SortColumnsForm : Form
    {

        private List<string> _allColumnsList;
        private Form1 _parentForm;
        private List<DataGridViewColumn> _dgvColumnList;
        private const string SORT_MODE_ASCENDING = "Ascending";
        private const string SORT_MODE_DESCENDING = "Descending";
        
        /// <summary>
        /// Constructorfor DataTableControl SortColumns
        /// </summary>
        /// <param name="dataTableControl"></param>
        /// <param name="dgvColumnList"></param>
        /// <param name="sortOrderList"></param>
        public SortColumnsForm(Form1 form, List<DataGridViewColumn> dgvColumnList)
        {
            InitializeComponent();
            _dgvColumnList = dgvColumnList;
            _parentForm = form;

            //Initialize datasource for column field
            _allColumnsList = new List<string>();
            foreach (DataGridViewColumn dgvCol in dgvColumnList)
                _allColumnsList.Add(dgvCol.HeaderText);

            //Load Column List
            cmbField.DataSource = _allColumnsList;

        }

        private void SortColumnsForm_Load(object sender, EventArgs e)
        {
            //Add a default row while loading, if there is no sort order saved earlier
            if (dgvSortColumns.Rows.Count == 0)
                AddRow();
        }

        private void btnAddLevel_Click(object sender, EventArgs e)
        {
            AddRow();
        }

        /// <summary>
        /// Add a new Row
        /// </summary>
        private void AddRow()
        {
            DataGridViewRow dgvRow = new DataGridViewRow();
            if (dgvSortColumns.Rows.Count == 0)
                dgvRow.HeaderCell.Value = "Sort by";
            else
                dgvRow.HeaderCell.Value = "Then by";

            dgvSortColumns.Rows.Add(dgvRow);
            dgvSortColumns.Rows[dgvSortColumns.Rows.Count - 1].Cells[1].Value = SORT_MODE_ASCENDING;
        }

        private void btnDeleteLevel_Click(object sender, EventArgs e)
        {
            DataGridViewSelectedRowCollection dgvRowColl = dgvSortColumns.SelectedRows;

            if (dgvRowColl.Count == 0)
            {
                MessageBox.Show("Please select a row.", "Sort", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            foreach (DataGridViewRow dgvRow in dgvRowColl)
            {
                    dgvSortColumns.Rows.Remove(dgvRow);
            }

            //Set First row header cell value
            if (dgvSortColumns.Rows.Count > 0 && dgvSortColumns.Rows[0].HeaderCell.Value.ToString() != "Sort by")
                dgvSortColumns.Rows[0].HeaderCell.Value = "Sort by";
        }

        /// <summary>
        /// Returns the first repeated column value
        /// Returns empty if not repeated
        /// </summary>
        /// <returns>string.empty if not repeated</returns>
        private string GetRepeatedColumn()
        {
            List<string> tempColumnList = new List<string>();
            foreach (DataGridViewRow dgvRow in dgvSortColumns.Rows)
            {
                if (tempColumnList.Contains(dgvRow.Cells[0].Value.ToString()))
                    return dgvRow.Cells[0].Value.ToString();
                else
                    tempColumnList.Add(dgvRow.Cells[0].Value.ToString());
            }

            return string.Empty;
        }

        /// <summary>
        /// Returns true if there is a cell not assigned
        /// </summary>
        /// <returns></returns>
        private bool HaveUnAssignedCells()
        {
            foreach (DataGridViewRow dgvRow in dgvSortColumns.Rows)
                foreach (DataGridViewCell dgvCell in dgvRow.Cells)
                    if (Convert.ToString(dgvCell.Value) == string.Empty || dgvCell.Value == null)
                        return true;

            return false;
        }

        private void btnSort_Click(object sender, EventArgs e)
        {

            if (dgvSortColumns.Rows.Count == 0 && chkCloseSort.Checked)
            {
                if (MessageBox.Show("No Sort criteria defined, Please press 'Yes' to confirm.", "Sort", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.Yes)
                {
                    this.DialogResult = DialogResult.OK;
                    this.Close();
                }
                else
                    return;
            }
            else if (dgvSortColumns.Rows.Count == 0 && !chkCloseSort.Checked)
            {
                MessageBox.Show("No Sort criteria defined. Please complete the sort criteria and try again.", "Sort", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (HaveUnAssignedCells())
            {
                MessageBox.Show("Sort criteria contains unassigned columns. Please complete the sort criteria and try again.", "Sort", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (GetRepeatedColumn() != string.Empty)
            {
                MessageBox.Show("'" + GetRepeatedColumn() + "' is being sorted more than once. Delete the duplicate sort criteria and try again.", "Sort", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (chkCloseSort.Checked)
            {
                this.DialogResult = DialogResult.OK;
                this.Close();
            }
            else
            {
                _parentForm.PerformSort(SortOrderList);
            }
        }
        private void btnReset_Click(object sender, EventArgs e)
        {
            //get List of all Rows
            List<DataGridViewRow> dgvRows  = new List<DataGridViewRow>();
            foreach(DataGridViewRow dgvRow in dgvSortColumns.Rows)
                dgvRows.Add(dgvRow);

            //Delete the roows in the above list
            foreach (DataGridViewRow dgvRow in dgvRows)
                dgvSortColumns.Rows.Remove(dgvRow);
        }


        /// <summary>
        /// DataGridView ComboBox Single Click Drop Down Items
        /// </summary>
        private void dataGridViewSortColumns_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            //exit if header column clicked
            if ((e.ColumnIndex == -1) || (e.RowIndex == -1))
                return;

            // the combobox column index
            if (e.ColumnIndex == 0 || e.ColumnIndex == 1)
            {
                dgvSortColumns.BeginEdit(false);

                if (this.dgvSortColumns.EditingControl != null && this.dgvSortColumns.EditingControl is ComboBox)
                {
                    ComboBox cmb = this.dgvSortColumns.EditingControl as ComboBox;
                    cmb.DroppedDown = true;
                }
            }
        }

        /// <summary>
        /// Get - List of New Sort Order
        /// </summary>
        internal List<KeyValuePair<DataGridViewColumn, bool>> SortOrderList
        {
            get
            {
                List<KeyValuePair<DataGridViewColumn, bool>> sortList = new List<KeyValuePair<DataGridViewColumn, bool>>();
                foreach (DataGridViewRow dgvRow in dgvSortColumns.Rows)
                {
                    //get 1.ColumnNmae, 2.SortMode
                    DataGridViewColumn myCol = null;
                    foreach (DataGridViewColumn dgvCol in _dgvColumnList)
                        if (dgvCol.HeaderText == dgvRow.Cells[0].Value.ToString())
                        {
                            myCol = dgvCol;
                            break;
                        }

                    bool isAscending = dgvRow.Cells[1].Value.ToString() == SORT_MODE_ASCENDING ? true : false;

                    sortList.Add(new KeyValuePair<DataGridViewColumn, bool>(myCol, isAscending));
                }

                return sortList;
            }
        }
    }
}
