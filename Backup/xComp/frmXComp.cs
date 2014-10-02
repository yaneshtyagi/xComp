using System;
using System.Collections.Generic;
using System.Collections;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace yCompnents.OfficeTools.xComp
{
    public partial class frmXComp : Form
    {
        ArrayList alDiff = new ArrayList(); //Holds references to all differences
        ArrayList alChanged = new ArrayList();
        int currDiff = -1; // Current difference count
        Hashtable htIgnoreDifferenceList = new Hashtable();
        bool MatchCase = false;
        bool ShowIdenticalRows = true;
        bool SyncScrool = true;

        public frmXComp()
        {
            InitializeComponent();
        }

        private void frmXComp_Load(object sender, EventArgs e)
        {
            //htIgnoreDifferenceList.Add("Marketing", "TV");

            toolStripStatusLabel3.Text = MatchCase? "Match Case": "Ignore Case";

            splitContainer1.SplitterDistance = splitContainer1.Width / 2; 
            splitContainer1.Width = panel2.Width - 10;
            splitContainer1.Height = panel2.Height - 30;
            if (string.IsNullOrEmpty(this.LeftFileName) || string.IsNullOrEmpty(this.RightFileName))
            {
                txtLeftFileName.Text = @"D:\Prod_Config\ProdKeys\ProdKeys_FFH.xlsx";
                txtRightFileName.Text = @"D:\Prod_Config\ProdKeys\ProdKeys_TV.xlsx";
            }
            else
            {
                txtLeftFileName.Text = this.LeftFileName;
                txtRightFileName.Text = this.RightFileName;
                btnCompare_Click(null, null);
            }
        }

        private DataSet loadExcelSheet(string FileName)
        {
            //loads excel sheet into DataReader
            string excelConStr = string.Format(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties=Excel 12.0", FileName);
            using (OleDbConnection excelCon = new OleDbConnection(excelConStr))
            {
                excelCon.Open();
                OleDbCommand cmd = new OleDbCommand("SELECT * FROM [Sheet1$]", excelCon);
                DataSet ds = new DataSet();
                OleDbDataAdapter da = new OleDbDataAdapter(cmd);
                da.Fill(ds, "Config");
                return ds;
            }

        }

        private void compareGrids(DataGridView LeftGrid, DataGridView RightGrid)
        {
            //Clear Difference Table
            alDiff.Clear();
            //compare Grids and set colors
            int maxRows, minRows;
            int maxCols, minCols;
            maxRows = (dgvLeft.RowCount >= dgvRight.RowCount) ? dgvLeft.RowCount : dgvRight.RowCount;
            maxCols = (dgvLeft.ColumnCount >= dgvRight.ColumnCount) ? dgvLeft.ColumnCount : dgvRight.ColumnCount;
            minRows= (dgvLeft.RowCount <= dgvRight.RowCount) ? dgvLeft.RowCount : dgvRight.RowCount;
            minCols = (dgvLeft.ColumnCount <= dgvRight.ColumnCount) ? dgvLeft.ColumnCount : dgvRight.ColumnCount;
            for (int i = 0; i < minRows; i++)
            {
                bool identical = true;
                for (int j = 0; j < minCols; j++)
                {
                    if (dgvLeft.Rows[i].Cells[j].Value != null || dgvRight.Rows[i].Cells[j].Value != null)
                    {
                        if (!CompareValues(dgvLeft.Rows[i].Cells[j].Value.ToString(), dgvRight.Rows[i].Cells[j].Value.ToString()))
                        {
                            dgvLeft.Rows[i].DefaultCellStyle.BackColor = Color.Pink;
                            dgvRight.Rows[i].DefaultCellStyle.BackColor = Color.Pink;
                            dgvLeft.Rows[i].Cells[j].Style.BackColor = Color.Red;
                            dgvRight.Rows[i].Cells[j].Style.BackColor = Color.Red;
                            alDiff.Add(new Difference(i, j));
                            identical = false;
                        }
                    }
                }
                if (identical && i >= 0 && !ShowIdenticalRows)
                {
                    dgvLeft.Rows[i].Visible = false;
                    dgvRight.Rows[i].Visible = false;
                }
            }
        }

        private bool CompareValues(string Value1, string Value2)
        {

            Hashtable htTemp = new Hashtable();
            if (!MatchCase)
            {
                Value1 = Value1.ToLower();
                Value2 = Value2.ToLower();
                foreach (string s in htIgnoreDifferenceList.Keys)
                    htTemp.Add(s.ToLower(), htIgnoreDifferenceList[s].ToString().ToLower());
            }

            foreach (string s in htTemp.Keys)
                if (Value1.Contains(s))
                    Value1 = Value1.Replace(s, htTemp[s].ToString());

            foreach (string s in htTemp.Keys)
                if (Value2.Contains(s))
                    Value2 = Value1.Replace(s, htTemp[s].ToString());

            return Value1.Equals(Value2);

        }

        private void btnFileBrowse_Click(object sender, EventArgs e)
        {
            openFileDialog1.Filter = "Excel 2007 Sheet|*.xlsx| Excel Old Version Sheets|*.xls| Macro Enables Excel Sheet|*.xlsm";
            openFileDialog1.Multiselect = false;
            openFileDialog1.ShowDialog();

            switch ((sender as Button).Name)
            {
                case "btnLeftFileBrowse":
                    txtLeftFileName.Text = openFileDialog1.FileName;
                    break;
                case "btnRightFileBrowse":
                    txtRightFileName.Text = openFileDialog1.FileName;
                    break;
            }


        }

        private void btnCompare_Click(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(txtLeftFileName.Text))
            {
                dgvLeft.DataSource = loadExcelSheet(txtLeftFileName.Text).Tables[0];
                dgvLeft.AutoGenerateColumns = true;
                
            }


            if (!string.IsNullOrEmpty(txtRightFileName.Text))
            {
                dgvRight.DataSource = loadExcelSheet(txtRightFileName.Text).Tables[0];
                dgvLeft.AutoGenerateColumns = true;
            }

            compareGrids(dgvLeft, dgvRight);
            toolStripStatusLabel1.Text = string.Format("{0} Differences", alDiff.Count);

        }

        private void dgv_Scroll(DataGridView sender, int row)
        {

            sender.FirstDisplayedScrollingRowIndex = row;
        }

        private void btnNextDiff_Click(object sender, EventArgs e)
        {
            if (currDiff < alDiff.Count)
            {
                for(int i = 0; i < dgvLeft.RowCount; i++)
                    for (int j = 0; j < dgvLeft.ColumnCount; j++)
                        dgvLeft.Rows[i].Cells[j].Selected = false;

                for (int i = 0; i < dgvRight.RowCount; i++)
                    for (int j = 0; j < dgvRight.ColumnCount; j++)
                        dgvRight.Rows[i].Cells[j].Selected = false;

                currDiff++;
                dgvLeft.Rows[(alDiff[currDiff] as Difference).row].Cells[(alDiff[currDiff] as Difference).col].Selected = true;
                dgvRight.Rows[(alDiff[currDiff] as Difference).row].Cells[(alDiff[currDiff] as Difference).col].Selected = true;
                ScrollViews();
            }
        }

        private void btnPrevDiff_Click(object sender, EventArgs e)
        {
            if (currDiff >= 0)
            {
                //Clear all selected rows from Left Grid
                for (int i = 0; i < dgvLeft.RowCount; i++)
                    for (int j = 0; j < dgvLeft.ColumnCount; j++)
                        dgvLeft.Rows[i].Cells[j].Selected = false;

                //Clear all selected rows from right grid
                for (int i = 0; i < dgvRight.RowCount; i++)
                    for (int j = 0; j < dgvRight.ColumnCount; j++)
                        dgvRight.Rows[i].Cells[j].Selected = false;

                if (currDiff > 0) currDiff--;
                dgvLeft.Rows[(alDiff[currDiff] as Difference).row].Cells[(alDiff[currDiff] as Difference).col].Selected = true;
                dgvRight.Rows[(alDiff[currDiff] as Difference).row].Cells[(alDiff[currDiff] as Difference).col].Selected = true;

                ScrollViews();
                
            }
        }

        private void ScrollViews()
        {
            int rowToScroll = (alDiff[currDiff] as Difference).row - 7;
            if (rowToScroll < 0) rowToScroll = 0;
            dgv_Scroll(dgvLeft, rowToScroll);
            dgv_Scroll(dgvRight, rowToScroll);
        }

        private void toolStripSplitButton1_ButtonClick(object sender, EventArgs e)
        {

        }

        private void toolStripComboBox1_Click(object sender, EventArgs e)
        {

        }

        private void statusStrip1_ItemClicked(object sender, ToolStripItemClickedEventArgs e)
        {

        }

        private void matchCaseToolStripMenuItem_Click(object sender, EventArgs e)
        {
            toolStripStatusLabel3.Text = "Match Case";
            MatchCase = true;
        }

        private void ignoreCaseToolStripMenuItem_Click(object sender, EventArgs e)
        {
            toolStripStatusLabel3.Text = "Ignore Case";
            MatchCase = false;
        }

        private void lbtnAbout_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            Form frmAbout = new About();
            frmAbout.ShowDialog();
        }

        private void btnExit_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void dgvRight_Scroll(object sender, ScrollEventArgs e)
        {
            if (SyncScrool) dgvLeft.FirstDisplayedScrollingRowIndex = dgvRight.FirstDisplayedScrollingRowIndex;
            if (SyncScrool) dgvLeft.FirstDisplayedScrollingColumnIndex = dgvRight.FirstDisplayedScrollingColumnIndex;
        }

        private void dgvLeft_Scroll(object sender, ScrollEventArgs e)
        {
            if (SyncScrool) dgvRight.FirstDisplayedScrollingRowIndex = dgvLeft.FirstDisplayedScrollingRowIndex;
            if (SyncScrool) dgvRight.FirstDisplayedScrollingColumnIndex = dgvLeft.FirstDisplayedScrollingColumnIndex;
        }

        private void btnMergeFromLeftToRight_Click(object sender, EventArgs e)
        {
            int row = dgvLeft.SelectedCells[0].RowIndex;
            int col = dgvLeft.SelectedCells[0].ColumnIndex;
            dgvRight.Rows[row].Cells[col].Value = dgvLeft.SelectedCells[0].Value;
            dgvRight.Rows[row].Cells[col].Style.BackColor = Color.Green;
            alChanged.Add(new Difference(row, col));
            //Write code to save excelsheet on clicking save button
        }

        private void UpdateSiteFile()
        {
            string excelConStr = string.Format(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties=Excel 12.0", txtRightFileName.Text);
            for (int i = 0; i < dgvRight.Rows.Count; i++)
            {
                using (OleDbConnection excelCon = new OleDbConnection(excelConStr))
                {
                    excelCon.Open();
                    foreach (Difference d in alChanged)
                    {
                        string keyColName = "ID";
                        string keyColValue = dgvRight.Rows[d.row].Cells[2].Value.ToString();
                        string colName = dgvRight.Columns[d.col].HeaderText;
                        string colValue = dgvRight.Rows[d.row].Cells[d.col].Value.ToString();

                        string strCmd = string.Format("UPDATE [Sheet1$] Set [{0}] = \"{1}\" where [{2}] = \"{3}\"", colName, colValue, keyColName, keyColValue);
                        OleDbCommand cmd = new OleDbCommand(strCmd, excelCon);
                        cmd.ExecuteNonQuery();
                        cmd = null;
                    }
                }
                //progressBar1.Value = i;
            }
        }

        private void btnSaveRightFile_Click(object sender, EventArgs e)
        {
            UpdateSiteFile();
        }

    }

    public class Difference
    {
        private int _row;
        private int _col;
        public int row
        {
            get
            {
                return _row;
            }
            set
            {
                _row = value;
            }
        }
        public int col
        {
            get
            {
                return _col;
            }
            set
            {
                _col = value;
            }
        }

        public Difference(int row, int col)
        {
            this.row = row;
            this.col = col;
        }
    }

}
