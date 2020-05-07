using System;
using System.Collections.Generic;
using System.Collections;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Linq;
using ExcelDataReader;
using OfficeOpenXml;

namespace yCompnents.OfficeTools.xComp
{
    public partial class frmXComp : Form
    {
        #region private fields
        ArrayList alDiff = new ArrayList(); //Holds references to all differences
        ArrayList alChanged = new ArrayList();
        int currDiff = -1; // Current difference count
        Hashtable htIgnoreDifferenceList = new Hashtable();
        bool MatchCase = false;
        bool ShowIdenticalRows = true;
        bool SyncScrool = true;
        bool clearGridOnReload = false;
        bool withinToleranceRange = false;
        #endregion

        public frmXComp()
        {
            InitializeComponent();
        }

        private void frmXComp_Load(object sender, EventArgs e)
        {
            try
            {
                //htIgnoreDifferenceList.Add("Marketing", "TV");

                toolStripStatusLabel3.Text = MatchCase ? "Match Case" : "Ignore Case";

                splitContainer1.SplitterDistance = splitContainer1.Width / 2;
                splitContainer1.Width = panel2.Width - 10;
                splitContainer1.Height = panel2.Height - 30;
                if (string.IsNullOrEmpty(this.LeftFileName) || string.IsNullOrEmpty(this.RightFileName))
                {
#if (DEBUG)
                    //Default values for testing
                    txtLeftFileName.Text = @"E:\Yanesh\Work\Projects\Current Project\Optum\RnD\xcomp comp\ComMod11innet.xlsm";
                    //txtRightFileName.Text = @"E:\Yanesh\Work\Projects\Current Project\Optum\RnD\xcomp comp\ComMod11innet_2.xlsm";
                    txtRightFileName.Text = @"E:\Yanesh\Work\Projects\Current Project\Optum\RnD\xcomp comp\g1.xlsx";
                    txtLeftSheet.Text = "Header";
                    txtRightSheet.Text = "Header";
                    txtLeftRange.Text = "A18:E56";
                    txtRightRange.Text = "B2:F40";
#endif
                }
                else
                {
                    txtLeftFileName.Text = this.LeftFileName;
                    txtRightFileName.Text = this.RightFileName;
                    btnCompare_Click(null, null);
                }
            }
            catch { }
        }

        private DataSet loadExcelSheet(string FileName)
        {
            try
            {
                toolStripProgressBar1.Value += 5;
                //loads excel sheet into DataReader
                string excelConStr = string.Format(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties=Excel 12.0", FileName);
                using (OleDbConnection excelCon = new OleDbConnection(excelConStr))
                {
                    excelCon.Open();
                    OleDbCommand cmd = new OleDbCommand("SELECT * FROM [Header$]", excelCon);
                    DataSet ds = new DataSet();
                    OleDbDataAdapter da = new OleDbDataAdapter(cmd);
                    da.Fill(ds, "Config");
                    return ds;
                }
            }
            catch { return null; }
        }


        private DataTable loadExcelSheet1(string fileName, string sheet, string range)
        {
            try
            {
                toolStripProgressBar1.Value += 5;
                FileInfo existingFile = new FileInfo(fileName);
                using (ExcelPackage package = new ExcelPackage((existingFile)))
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets[sheet];
                    ExcelRangeBase erb = worksheet.Cells[range];

                    DataTable dt = new DataTable();
                    for (int i = erb.Start.Column; i <= erb.End.Column; i++)
                        dt.Columns.Add(i.ToString());

                    for (int row = erb.Start.Row; row <= erb.End.Row; row++)
                    {
                        DataRow dr = dt.NewRow();
                        int x = 0;
                        for (int col = erb.Start.Column; col <= erb.End.Column; col++)
                        {
                            dr[x++] = worksheet.Cells[row, col].Value;
                        }
                        dt.Rows.Add(dr);
                    }
                    return dt;
                }
            }
            catch { return null; }
        }

        int GetColumnNumber(string name)
        {
            try
            {
                int number = 0;
                int pow = 1;
                for (int i = name.Length - 1; i >= 0; i--)
                {
                    number += (name[i] - 'A' + 1) * pow;
                    pow *= 26;
                }

                return number;
            }
            catch
            {
                throw;
            }
            //return name.ToUpper().Aggregate(0, (column, letter) => 26 * column + letter - 'A' + 1);
        }

        Int32 GetColumnNumber1(string name, out int col)
        {
            col = 0;
            name = name.ToUpper();
            char[] chars = new char[name.Length];
            string row = "";
            int i = 0;
            foreach (char c in name)
            {
                if (char.IsLetter(c))
                {
                    col += c - 65;
                }
                else
                {
                    row += c.ToString();
                }
            }

            return Int32.Parse(row);

            //return name.ToUpper().Aggregate(0, (column, letter) => 26 * column + letter - 'A' + 1);
        }
        private DataSet loadExcelSheet(string FileName, string sheetName)
        {
            if (string.IsNullOrEmpty(sheetName))
            {
                return null;
            }

            //loads selected excel sheet into DataReader
            string excelConStr = string.Format(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties=Excel 12.0", FileName);
            using (OleDbConnection excelCon = new OleDbConnection(excelConStr))
            {
                excelCon.Open();
                OleDbCommand cmd = new OleDbCommand("SELECT * FROM [" + sheetName + "$]", excelCon);
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
            minRows = (dgvLeft.RowCount <= dgvRight.RowCount) ? dgvLeft.RowCount : dgvRight.RowCount;
            minCols = (dgvLeft.ColumnCount <= dgvRight.ColumnCount) ? dgvLeft.ColumnCount : dgvRight.ColumnCount;
            toolStripProgressBar1.Maximum = (minRows * minCols) + 15;
            for (int i = 0; i < minRows; i++)
            {
                bool identical = true;
                for (int j = 0; j < minCols; j++)
                {
                    Application.DoEvents();
                    toolStripProgressBar1.Value += 1;
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
                        else
                        {
                            //if values are not equal but within tolerance range, highlight with blue.
                            if (withinToleranceRange)
                            {
                                dgvLeft.Rows[i].DefaultCellStyle.BackColor = Color.LightBlue;
                                dgvRight.Rows[i].DefaultCellStyle.BackColor = Color.LightBlue;
                                dgvLeft.Rows[i].Cells[j].Style.BackColor = Color.Cyan;
                                dgvRight.Rows[i].Cells[j].Style.BackColor = Color.Cyan;
                            }
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

            return CompareValues(Value1, Value2, !string.IsNullOrEmpty(txtTolerance.Text));

            //below code has been depricated having logic being implemented in overridden method.
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

        private bool CompareValues(string Value1, string Value2, bool useTolerance)
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

            double numValue1, numValue2;
            if (!useTolerance || string.IsNullOrEmpty(txtTolerance.Text) || !double.TryParse(Value1, out numValue1) || !double.TryParse(Value2, out numValue2))
                return Value1.Equals(Value2);



            //implement tollrance (for numeric values only)
            double leftValue = 0.0;
            double rightValue = 0.0;
            double tolerance = 0.0;
            var errorMessage = "";

            if (!double.TryParse(txtTolerance.Text.ToString(), out tolerance))
                errorMessage = "Invalid Left Value";

            if (!double.TryParse(Value1, out leftValue))
                errorMessage = "Invalid Right Value";

            if (!double.TryParse(Value2, out rightValue))
                errorMessage = "Invalid Tolerance";


            if (!string.IsNullOrEmpty(errorMessage))
            {
                //MessageBox.Show(errorMessage);
                return false;
            }
            if (Math.Abs(leftValue - rightValue) <= tolerance)
            {
                //if values are not equal but within the tolerance range, set the withinToleranceRange field to true.
                withinToleranceRange = (Math.Abs(leftValue - rightValue) != 0.00);
                return true;
            }
            else
                return false;

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
            //clear the grids before reload. This will give user a visual reload experience.
            if (clearGridOnReload)
            {
                dgvLeft.DataSource = null;
                dgvRight.DataSource = null;
            }

            if (!string.IsNullOrEmpty(txtLeftFileName.Text))
            {
                if (string.IsNullOrEmpty(txtLeftRange.Text))
                {
                    //dgvLeft.DataSource = loadExcelSheet(txtLeftFileName.Text).Tables[0];
                    //dgvLeft.DataSource = loadExcelSheet1(txtLeftFileName.Text, txtLeftSheet.Text, txtLeftRange.Text);
                    dgvLeft.DataSource = loadExcelSheet2(txtLeftFileName.Text, txtLeftSheet.Text, txtLeftRange.Text);
                }
                else
                {
                    //dgvLeft.DataSource = GetExcelRange(txtLeftRange.Text);
                    //dgvLeft.DataSource = loadExcelSheet1(txtLeftFileName.Text, txtLeftSheet.Text, txtLeftRange.Text);
                    dgvLeft.DataSource = loadExcelSheet2(txtLeftFileName.Text, txtLeftSheet.Text, txtLeftRange.Text);
                }
                dgvLeft.AutoGenerateColumns = true;
            }


            if (!string.IsNullOrEmpty(txtRightFileName.Text))
            {
                if (string.IsNullOrEmpty(txtRightRange.Text))
                {
                    //dgvRight.DataSource = loadExcelSheet(txtRightFileName.Text).Tables[0];
                    dgvRight.DataSource = loadExcelSheet2(txtRightFileName.Text, txtRightSheet.Text, txtRightRange.Text);
                }
                else
                {
                    //dgvRight.DataSource = GetExcelRange(txtRightRange.Text);
                    dgvRight.DataSource = loadExcelSheet2(txtRightFileName.Text, txtRightSheet.Text, txtRightRange.Text);
                }
                dgvLeft.AutoGenerateColumns = true;
            }

            compareGrids(dgvLeft, dgvRight);
            toolStripStatusLabel1.Text = string.Format("{0} Differences", alDiff.Count);

        }

        //Reads a datatable and retruns the specified rowcolumn Range
        //Range format A17:E39
        private DataTable GetExcelRange(string range)
        {
            string[] rangeArray = range.Split(':');

            int beginCol, endCol;
            var beginRow = GetColumnNumber1(rangeArray[0], out beginCol);
            var endRow = GetColumnNumber1(rangeArray[1], out endCol);

            var dt = loadExcelSheet(txtLeftFileName.Text).Tables[0];

            var dataTable = new DataTable();
            for (int i = beginCol; i < endCol; i++)
                dataTable.Columns.Add(i.ToString());

            for (int i = beginRow; i < endRow; i++)
            {
                DataRow dr = dataTable.NewRow();
                int x = 0;
                for (int j = beginCol; j < endCol; j++)
                {
                    dr[x] = dt.Rows[i][j];
                    x++;
                }
                dataTable.Rows.Add(dr);
            }
            return dataTable;
        }

        private void dgv_Scroll(DataGridView sender, int row)
        {

            sender.FirstDisplayedScrollingRowIndex = row;
        }

        private void btnNextDiff_Click(object sender, EventArgs e)
        {
            try
            {
                if (alDiff.Count == 0) return;
                if (currDiff < alDiff.Count)
                {
                    for (int i = 0; i < dgvLeft.RowCount; i++)
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
            catch { }
        }

        private void btnPrevDiff_Click(object sender, EventArgs e)
        {
            try
            {
                if (alDiff.Count == 0) return;
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
            catch { }
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

        private void doNotClearGridOnReloadToolStripMenuItem_Click(object sender, EventArgs e)
        {
            DoNotClearGridOnReloadLabel.Text = "Do not clear grid on reload";
            this.clearGridOnReload = false;
        }

        private void clareGridOnReloadToolStripMenuItem_Click(object sender, EventArgs e)
        {
            DoNotClearGridOnReloadLabel.Text = "Clear grid on reload";
            this.clearGridOnReload = true;
        }

        private DataTable loadExcelSheet2(string fileName, string sheet, string range)
        {
            try
            {
                toolStripProgressBar1.Value += 5;
                FileInfo existingFile = new FileInfo(fileName);


                using (var stream = File.Open(fileName, FileMode.Open, FileAccess.Read))
                {
                    using (var reader = ExcelReaderFactory.CreateReader(stream))
                    {
                        
                        var x = reader.AsDataSet(new ExcelDataSetConfiguration()
                        {
                            UseColumnDataType = true,
                            FilterSheet = (tableReader, sheetIndex) =>  tableReader.Name.Equals(sheet)
                        });
                        return x.Tables[0];
                    }
                }
            }
            catch { return null; }
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
