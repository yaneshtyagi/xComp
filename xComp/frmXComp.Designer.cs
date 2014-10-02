namespace yCompnents.OfficeTools.xComp
{
    partial class frmXComp
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmXComp));
            this.panel1 = new System.Windows.Forms.Panel();
            this.label6 = new System.Windows.Forms.Label();
            this.txtRightRange = new System.Windows.Forms.TextBox();
            this.label7 = new System.Windows.Forms.Label();
            this.txtLeftRange = new System.Windows.Forms.TextBox();
            this.label5 = new System.Windows.Forms.Label();
            this.txtRightSheet = new System.Windows.Forms.TextBox();
            this.label4 = new System.Windows.Forms.Label();
            this.txtLeftSheet = new System.Windows.Forms.TextBox();
            this.txtTolerance = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.numericUpDown1 = new System.Windows.Forms.NumericUpDown();
            this.lblDecimalPlace = new System.Windows.Forms.Label();
            this.btnSaveRightFile = new System.Windows.Forms.Button();
            this.btnMergeFromLeftToRight = new System.Windows.Forms.Button();
            this.btnExit = new System.Windows.Forms.Button();
            this.lbtnAbout = new System.Windows.Forms.LinkLabel();
            this.btnPrevDiff = new System.Windows.Forms.Button();
            this.btnNextDiff = new System.Windows.Forms.Button();
            this.btnCompare = new System.Windows.Forms.Button();
            this.btnRightFileBrowse = new System.Windows.Forms.Button();
            this.txtRightFileName = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.btnLeftFileBrowse = new System.Windows.Forms.Button();
            this.txtLeftFileName = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.panel2 = new System.Windows.Forms.Panel();
            this.splitContainer1 = new System.Windows.Forms.SplitContainer();
            this.dgvLeft = new System.Windows.Forms.DataGridView();
            this.dgvRight = new System.Windows.Forms.DataGridView();
            this.panel4 = new System.Windows.Forms.Panel();
            this.statusStrip1 = new System.Windows.Forms.StatusStrip();
            this.toolStripProgressBar1 = new System.Windows.Forms.ToolStripProgressBar();
            this.toolStripStatusLabel1 = new System.Windows.Forms.ToolStripStatusLabel();
            this.toolStripStatusLabel2 = new System.Windows.Forms.ToolStripStatusLabel();
            this.toolStripDropDownButton1 = new System.Windows.Forms.ToolStripDropDownButton();
            this.matchCaseToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.ignoreCaseToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripStatusLabel3 = new System.Windows.Forms.ToolStripStatusLabel();
            this.toolStripDropDownButton2 = new System.Windows.Forms.ToolStripDropDownButton();
            this.doNotClearGridOnReloadToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.clareGridOnReloadToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.DoNotClearGridOnReloadLabel = new System.Windows.Forms.ToolStripStatusLabel();
            this.panel1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.numericUpDown1)).BeginInit();
            this.panel2.SuspendLayout();
            this.splitContainer1.Panel1.SuspendLayout();
            this.splitContainer1.Panel2.SuspendLayout();
            this.splitContainer1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgvLeft)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.dgvRight)).BeginInit();
            this.panel4.SuspendLayout();
            this.statusStrip1.SuspendLayout();
            this.SuspendLayout();
            // 
            // panel1
            // 
            this.panel1.Controls.Add(this.label6);
            this.panel1.Controls.Add(this.txtRightRange);
            this.panel1.Controls.Add(this.label7);
            this.panel1.Controls.Add(this.txtLeftRange);
            this.panel1.Controls.Add(this.label5);
            this.panel1.Controls.Add(this.txtRightSheet);
            this.panel1.Controls.Add(this.label4);
            this.panel1.Controls.Add(this.txtLeftSheet);
            this.panel1.Controls.Add(this.txtTolerance);
            this.panel1.Controls.Add(this.label3);
            this.panel1.Controls.Add(this.numericUpDown1);
            this.panel1.Controls.Add(this.lblDecimalPlace);
            this.panel1.Controls.Add(this.btnSaveRightFile);
            this.panel1.Controls.Add(this.btnMergeFromLeftToRight);
            this.panel1.Controls.Add(this.btnExit);
            this.panel1.Controls.Add(this.lbtnAbout);
            this.panel1.Controls.Add(this.btnPrevDiff);
            this.panel1.Controls.Add(this.btnNextDiff);
            this.panel1.Controls.Add(this.btnCompare);
            this.panel1.Controls.Add(this.btnRightFileBrowse);
            this.panel1.Controls.Add(this.txtRightFileName);
            this.panel1.Controls.Add(this.label2);
            this.panel1.Controls.Add(this.btnLeftFileBrowse);
            this.panel1.Controls.Add(this.txtLeftFileName);
            this.panel1.Controls.Add(this.label1);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Top;
            this.panel1.Location = new System.Drawing.Point(0, 0);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(1010, 62);
            this.panel1.TabIndex = 0;
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(468, 38);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(39, 13);
            this.label6.TabIndex = 24;
            this.label6.Text = "Range";
            // 
            // txtRightRange
            // 
            this.txtRightRange.Location = new System.Drawing.Point(514, 35);
            this.txtRightRange.Name = "txtRightRange";
            this.txtRightRange.Size = new System.Drawing.Size(100, 20);
            this.txtRightRange.TabIndex = 23;
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(468, 13);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(39, 13);
            this.label7.TabIndex = 22;
            this.label7.Text = "Range";
            // 
            // txtLeftRange
            // 
            this.txtLeftRange.Location = new System.Drawing.Point(514, 10);
            this.txtLeftRange.Name = "txtLeftRange";
            this.txtLeftRange.Size = new System.Drawing.Size(100, 20);
            this.txtLeftRange.TabIndex = 21;
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(297, 37);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(66, 13);
            this.label5.TabIndex = 20;
            this.label5.Text = "Sheet Name";
            // 
            // txtRightSheet
            // 
            this.txtRightSheet.Location = new System.Drawing.Point(366, 34);
            this.txtRightSheet.Name = "txtRightSheet";
            this.txtRightSheet.Size = new System.Drawing.Size(100, 20);
            this.txtRightSheet.TabIndex = 19;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(297, 12);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(66, 13);
            this.label4.TabIndex = 18;
            this.label4.Text = "Sheet Name";
            // 
            // txtLeftSheet
            // 
            this.txtLeftSheet.Location = new System.Drawing.Point(366, 9);
            this.txtLeftSheet.Name = "txtLeftSheet";
            this.txtLeftSheet.Size = new System.Drawing.Size(100, 20);
            this.txtLeftSheet.TabIndex = 17;
            // 
            // txtTolerance
            // 
            this.txtTolerance.Location = new System.Drawing.Point(678, 10);
            this.txtTolerance.Name = "txtTolerance";
            this.txtTolerance.Size = new System.Drawing.Size(76, 20);
            this.txtTolerance.TabIndex = 16;
            this.txtTolerance.Text = "0.00";
            this.txtTolerance.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(623, 12);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(55, 13);
            this.label3.TabIndex = 15;
            this.label3.Tag = "";
            this.label3.Text = "Tolerance";
            // 
            // numericUpDown1
            // 
            this.numericUpDown1.Location = new System.Drawing.Point(838, 11);
            this.numericUpDown1.Name = "numericUpDown1";
            this.numericUpDown1.Size = new System.Drawing.Size(33, 20);
            this.numericUpDown1.TabIndex = 14;
            // 
            // lblDecimalPlace
            // 
            this.lblDecimalPlace.AutoSize = true;
            this.lblDecimalPlace.Location = new System.Drawing.Point(754, 13);
            this.lblDecimalPlace.Name = "lblDecimalPlace";
            this.lblDecimalPlace.Size = new System.Drawing.Size(80, 13);
            this.lblDecimalPlace.TabIndex = 13;
            this.lblDecimalPlace.Tag = "";
            this.lblDecimalPlace.Text = "Decimal Places";
            // 
            // btnSaveRightFile
            // 
            this.btnSaveRightFile.Font = new System.Drawing.Font("Wingdings", 11.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(2)));
            this.btnSaveRightFile.Location = new System.Drawing.Point(697, 33);
            this.btnSaveRightFile.Name = "btnSaveRightFile";
            this.btnSaveRightFile.Size = new System.Drawing.Size(23, 23);
            this.btnSaveRightFile.TabIndex = 12;
            this.btnSaveRightFile.Text = "<";
            this.btnSaveRightFile.UseVisualStyleBackColor = true;
            this.btnSaveRightFile.Click += new System.EventHandler(this.btnSaveRightFile_Click);
            // 
            // btnMergeFromLeftToRight
            // 
            this.btnMergeFromLeftToRight.Font = new System.Drawing.Font("Wingdings", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(2)));
            this.btnMergeFromLeftToRight.Location = new System.Drawing.Point(673, 33);
            this.btnMergeFromLeftToRight.Name = "btnMergeFromLeftToRight";
            this.btnMergeFromLeftToRight.Size = new System.Drawing.Size(23, 23);
            this.btnMergeFromLeftToRight.TabIndex = 11;
            this.btnMergeFromLeftToRight.Text = "è";
            this.btnMergeFromLeftToRight.UseVisualStyleBackColor = true;
            this.btnMergeFromLeftToRight.Click += new System.EventHandler(this.btnMergeFromLeftToRight_Click);
            // 
            // btnExit
            // 
            this.btnExit.Location = new System.Drawing.Point(722, 33);
            this.btnExit.Name = "btnExit";
            this.btnExit.Size = new System.Drawing.Size(40, 23);
            this.btnExit.TabIndex = 10;
            this.btnExit.Text = "E&xit";
            this.btnExit.UseVisualStyleBackColor = true;
            this.btnExit.Click += new System.EventHandler(this.btnExit_Click);
            // 
            // lbtnAbout
            // 
            this.lbtnAbout.AutoSize = true;
            this.lbtnAbout.Location = new System.Drawing.Point(766, 39);
            this.lbtnAbout.Name = "lbtnAbout";
            this.lbtnAbout.Size = new System.Drawing.Size(13, 13);
            this.lbtnAbout.TabIndex = 9;
            this.lbtnAbout.TabStop = true;
            this.lbtnAbout.Text = "?";
            this.lbtnAbout.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.lbtnAbout_LinkClicked);
            // 
            // btnPrevDiff
            // 
            this.btnPrevDiff.Font = new System.Drawing.Font("Wingdings", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(2)));
            this.btnPrevDiff.Location = new System.Drawing.Point(649, 33);
            this.btnPrevDiff.Name = "btnPrevDiff";
            this.btnPrevDiff.Size = new System.Drawing.Size(23, 23);
            this.btnPrevDiff.TabIndex = 8;
            this.btnPrevDiff.Text = "é";
            this.btnPrevDiff.UseVisualStyleBackColor = true;
            this.btnPrevDiff.Click += new System.EventHandler(this.btnPrevDiff_Click);
            // 
            // btnNextDiff
            // 
            this.btnNextDiff.Font = new System.Drawing.Font("Wingdings", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(2)));
            this.btnNextDiff.Location = new System.Drawing.Point(626, 33);
            this.btnNextDiff.Name = "btnNextDiff";
            this.btnNextDiff.Size = new System.Drawing.Size(23, 23);
            this.btnNextDiff.TabIndex = 7;
            this.btnNextDiff.Text = "ê";
            this.btnNextDiff.UseVisualStyleBackColor = true;
            this.btnNextDiff.Click += new System.EventHandler(this.btnNextDiff_Click);
            // 
            // btnCompare
            // 
            this.btnCompare.Location = new System.Drawing.Point(878, 8);
            this.btnCompare.Name = "btnCompare";
            this.btnCompare.Size = new System.Drawing.Size(120, 46);
            this.btnCompare.TabIndex = 6;
            this.btnCompare.Text = "&Compare";
            this.btnCompare.UseVisualStyleBackColor = true;
            this.btnCompare.Click += new System.EventHandler(this.btnCompare_Click);
            // 
            // btnRightFileBrowse
            // 
            this.btnRightFileBrowse.Location = new System.Drawing.Point(272, 32);
            this.btnRightFileBrowse.Name = "btnRightFileBrowse";
            this.btnRightFileBrowse.Size = new System.Drawing.Size(25, 23);
            this.btnRightFileBrowse.TabIndex = 5;
            this.btnRightFileBrowse.Text = "...";
            this.btnRightFileBrowse.UseVisualStyleBackColor = true;
            this.btnRightFileBrowse.Click += new System.EventHandler(this.btnFileBrowse_Click);
            // 
            // txtRightFileName
            // 
            this.txtRightFileName.Location = new System.Drawing.Point(94, 34);
            this.txtRightFileName.Name = "txtRightFileName";
            this.txtRightFileName.Size = new System.Drawing.Size(175, 20);
            this.txtRightFileName.TabIndex = 4;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(3, 37);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(92, 13);
            this.label2.TabIndex = 3;
            this.label2.Text = "Right Excel Sheet";
            // 
            // btnLeftFileBrowse
            // 
            this.btnLeftFileBrowse.Location = new System.Drawing.Point(272, 7);
            this.btnLeftFileBrowse.Name = "btnLeftFileBrowse";
            this.btnLeftFileBrowse.Size = new System.Drawing.Size(25, 23);
            this.btnLeftFileBrowse.TabIndex = 2;
            this.btnLeftFileBrowse.Text = "...";
            this.btnLeftFileBrowse.UseVisualStyleBackColor = true;
            this.btnLeftFileBrowse.Click += new System.EventHandler(this.btnFileBrowse_Click);
            // 
            // txtLeftFileName
            // 
            this.txtLeftFileName.Location = new System.Drawing.Point(94, 9);
            this.txtLeftFileName.Name = "txtLeftFileName";
            this.txtLeftFileName.Size = new System.Drawing.Size(175, 20);
            this.txtLeftFileName.TabIndex = 1;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(3, 12);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(85, 13);
            this.label1.TabIndex = 0;
            this.label1.Text = "Left Excel Sheet";
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.FileName = "openFileDialog1";
            // 
            // panel2
            // 
            this.panel2.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.panel2.Controls.Add(this.splitContainer1);
            this.panel2.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel2.Location = new System.Drawing.Point(0, 62);
            this.panel2.Name = "panel2";
            this.panel2.Size = new System.Drawing.Size(1010, 202);
            this.panel2.TabIndex = 1;
            // 
            // splitContainer1
            // 
            this.splitContainer1.Location = new System.Drawing.Point(0, 0);
            this.splitContainer1.Name = "splitContainer1";
            // 
            // splitContainer1.Panel1
            // 
            this.splitContainer1.Panel1.AutoScroll = true;
            this.splitContainer1.Panel1.Controls.Add(this.dgvLeft);
            // 
            // splitContainer1.Panel2
            // 
            this.splitContainer1.Panel2.AutoScroll = true;
            this.splitContainer1.Panel2.Controls.Add(this.dgvRight);
            this.splitContainer1.Size = new System.Drawing.Size(782, 100);
            this.splitContainer1.SplitterDistance = 260;
            this.splitContainer1.TabIndex = 0;
            // 
            // dgvLeft
            // 
            this.dgvLeft.AllowUserToAddRows = false;
            this.dgvLeft.AllowUserToDeleteRows = false;
            this.dgvLeft.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells;
            this.dgvLeft.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgvLeft.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dgvLeft.Location = new System.Drawing.Point(0, 0);
            this.dgvLeft.Name = "dgvLeft";
            this.dgvLeft.Size = new System.Drawing.Size(260, 100);
            this.dgvLeft.TabIndex = 0;
            this.dgvLeft.Scroll += new System.Windows.Forms.ScrollEventHandler(this.dgvLeft_Scroll);
            // 
            // dgvRight
            // 
            this.dgvRight.AllowUserToAddRows = false;
            this.dgvRight.AllowUserToDeleteRows = false;
            this.dgvRight.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells;
            this.dgvRight.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgvRight.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dgvRight.Location = new System.Drawing.Point(0, 0);
            this.dgvRight.Name = "dgvRight";
            this.dgvRight.Size = new System.Drawing.Size(518, 100);
            this.dgvRight.TabIndex = 0;
            this.dgvRight.Scroll += new System.Windows.Forms.ScrollEventHandler(this.dgvRight_Scroll);
            // 
            // panel4
            // 
            this.panel4.Controls.Add(this.statusStrip1);
            this.panel4.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.panel4.Location = new System.Drawing.Point(0, 236);
            this.panel4.Name = "panel4";
            this.panel4.Size = new System.Drawing.Size(1010, 28);
            this.panel4.TabIndex = 3;
            // 
            // statusStrip1
            // 
            this.statusStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.toolStripProgressBar1,
            this.toolStripStatusLabel1,
            this.toolStripStatusLabel2,
            this.toolStripDropDownButton1,
            this.toolStripStatusLabel3,
            this.toolStripDropDownButton2,
            this.DoNotClearGridOnReloadLabel});
            this.statusStrip1.Location = new System.Drawing.Point(0, 6);
            this.statusStrip1.Name = "statusStrip1";
            this.statusStrip1.Size = new System.Drawing.Size(1010, 22);
            this.statusStrip1.TabIndex = 5;
            this.statusStrip1.Text = "statusStrip1";
            this.statusStrip1.ItemClicked += new System.Windows.Forms.ToolStripItemClickedEventHandler(this.statusStrip1_ItemClicked);
            // 
            // toolStripProgressBar1
            // 
            this.toolStripProgressBar1.Name = "toolStripProgressBar1";
            this.toolStripProgressBar1.Size = new System.Drawing.Size(100, 16);
            this.toolStripProgressBar1.Style = System.Windows.Forms.ProgressBarStyle.Continuous;
            // 
            // toolStripStatusLabel1
            // 
            this.toolStripStatusLabel1.Name = "toolStripStatusLabel1";
            this.toolStripStatusLabel1.Size = new System.Drawing.Size(0, 17);
            // 
            // toolStripStatusLabel2
            // 
            this.toolStripStatusLabel2.Name = "toolStripStatusLabel2";
            this.toolStripStatusLabel2.Size = new System.Drawing.Size(0, 17);
            // 
            // toolStripDropDownButton1
            // 
            this.toolStripDropDownButton1.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            this.toolStripDropDownButton1.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.matchCaseToolStripMenuItem,
            this.ignoreCaseToolStripMenuItem});
            this.toolStripDropDownButton1.Image = ((System.Drawing.Image)(resources.GetObject("toolStripDropDownButton1.Image")));
            this.toolStripDropDownButton1.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.toolStripDropDownButton1.Name = "toolStripDropDownButton1";
            this.toolStripDropDownButton1.Size = new System.Drawing.Size(29, 20);
            this.toolStripDropDownButton1.Text = "Case Sensitivity";
            // 
            // matchCaseToolStripMenuItem
            // 
            this.matchCaseToolStripMenuItem.Name = "matchCaseToolStripMenuItem";
            this.matchCaseToolStripMenuItem.Size = new System.Drawing.Size(136, 22);
            this.matchCaseToolStripMenuItem.Text = "Match Case";
            this.matchCaseToolStripMenuItem.Click += new System.EventHandler(this.matchCaseToolStripMenuItem_Click);
            // 
            // ignoreCaseToolStripMenuItem
            // 
            this.ignoreCaseToolStripMenuItem.Name = "ignoreCaseToolStripMenuItem";
            this.ignoreCaseToolStripMenuItem.Size = new System.Drawing.Size(136, 22);
            this.ignoreCaseToolStripMenuItem.Text = "Ignore Case";
            this.ignoreCaseToolStripMenuItem.Click += new System.EventHandler(this.ignoreCaseToolStripMenuItem_Click);
            // 
            // toolStripStatusLabel3
            // 
            this.toolStripStatusLabel3.Name = "toolStripStatusLabel3";
            this.toolStripStatusLabel3.Size = new System.Drawing.Size(118, 17);
            this.toolStripStatusLabel3.Text = "toolStripStatusLabel3";
            // 
            // toolStripDropDownButton2
            // 
            this.toolStripDropDownButton2.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            this.toolStripDropDownButton2.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.doNotClearGridOnReloadToolStripMenuItem,
            this.clareGridOnReloadToolStripMenuItem});
            this.toolStripDropDownButton2.Image = ((System.Drawing.Image)(resources.GetObject("toolStripDropDownButton2.Image")));
            this.toolStripDropDownButton2.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.toolStripDropDownButton2.Name = "toolStripDropDownButton2";
            this.toolStripDropDownButton2.Size = new System.Drawing.Size(29, 20);
            this.toolStripDropDownButton2.Text = "toolStripDropDownButton2";
            // 
            // doNotClearGridOnReloadToolStripMenuItem
            // 
            this.doNotClearGridOnReloadToolStripMenuItem.Name = "doNotClearGridOnReloadToolStripMenuItem";
            this.doNotClearGridOnReloadToolStripMenuItem.Size = new System.Drawing.Size(215, 22);
            this.doNotClearGridOnReloadToolStripMenuItem.Text = "Do not clear grid on reload";
            this.doNotClearGridOnReloadToolStripMenuItem.Click += new System.EventHandler(this.doNotClearGridOnReloadToolStripMenuItem_Click);
            // 
            // clareGridOnReloadToolStripMenuItem
            // 
            this.clareGridOnReloadToolStripMenuItem.Name = "clareGridOnReloadToolStripMenuItem";
            this.clareGridOnReloadToolStripMenuItem.Size = new System.Drawing.Size(215, 22);
            this.clareGridOnReloadToolStripMenuItem.Text = "Clare Grid on Reload";
            this.clareGridOnReloadToolStripMenuItem.Click += new System.EventHandler(this.clareGridOnReloadToolStripMenuItem_Click);
            // 
            // DoNotClearGridOnReloadLabel
            // 
            this.DoNotClearGridOnReloadLabel.Name = "DoNotClearGridOnReloadLabel";
            this.DoNotClearGridOnReloadLabel.Size = new System.Drawing.Size(0, 17);
            // 
            // frmXComp
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1010, 264);
            this.Controls.Add(this.panel4);
            this.Controls.Add(this.panel2);
            this.Controls.Add(this.panel1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "frmXComp";
            this.Text = "xComp - Find the Difference in Excel";
            this.WindowState = System.Windows.Forms.FormWindowState.Maximized;
            this.Load += new System.EventHandler(this.frmXComp_Load);
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.numericUpDown1)).EndInit();
            this.panel2.ResumeLayout(false);
            this.splitContainer1.Panel1.ResumeLayout(false);
            this.splitContainer1.Panel2.ResumeLayout(false);
            this.splitContainer1.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dgvLeft)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.dgvRight)).EndInit();
            this.panel4.ResumeLayout(false);
            this.panel4.PerformLayout();
            this.statusStrip1.ResumeLayout(false);
            this.statusStrip1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        #region "Custom Properties"
        private string _LeftFileName;

        public string LeftFileName
        {
            get { return _LeftFileName; }
            set { _LeftFileName = value; }
        }

        private string _RightFileName;

        public string RightFileName
        {
            get { return _RightFileName; }
            set { _RightFileName = value; }
        }

        public frmXComp(string LeftFileName, string RightFileName)
        {
            this.LeftFileName = LeftFileName;
            this.RightFileName = RightFileName;
            InitializeComponent();
        }
	
        #endregion

        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Button btnLeftFileBrowse;
        private System.Windows.Forms.TextBox txtLeftFileName;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.OpenFileDialog openFileDialog1;
        private System.Windows.Forms.Button btnRightFileBrowse;
        private System.Windows.Forms.TextBox txtRightFileName;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Panel panel2;
        private System.Windows.Forms.SplitContainer splitContainer1;
        private System.Windows.Forms.DataGridView dgvLeft;
        private System.Windows.Forms.DataGridView dgvRight;
        private System.Windows.Forms.Panel panel4;
        private System.Windows.Forms.StatusStrip statusStrip1;
        private System.Windows.Forms.ToolStripProgressBar toolStripProgressBar1;
        private System.Windows.Forms.ToolStripStatusLabel toolStripStatusLabel1;
        private System.Windows.Forms.Button btnCompare;
        private System.Windows.Forms.Button btnNextDiff;
        private System.Windows.Forms.Button btnPrevDiff;
        private System.Windows.Forms.ToolStripStatusLabel toolStripStatusLabel2;
        private System.Windows.Forms.ToolStripDropDownButton toolStripDropDownButton1;
        private System.Windows.Forms.ToolStripMenuItem ignoreCaseToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem matchCaseToolStripMenuItem;
        private System.Windows.Forms.ToolStripStatusLabel toolStripStatusLabel3;
        private System.Windows.Forms.LinkLabel lbtnAbout;
        private System.Windows.Forms.Button btnExit;
        private System.Windows.Forms.Button btnMergeFromLeftToRight;
        private System.Windows.Forms.Button btnSaveRightFile;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.TextBox txtRightRange;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.TextBox txtLeftRange;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.TextBox txtRightSheet;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.TextBox txtLeftSheet;
        private System.Windows.Forms.TextBox txtTolerance;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.NumericUpDown numericUpDown1;
        private System.Windows.Forms.Label lblDecimalPlace;
        private System.Windows.Forms.ToolStripDropDownButton toolStripDropDownButton2;
        private System.Windows.Forms.ToolStripMenuItem doNotClearGridOnReloadToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem clareGridOnReloadToolStripMenuItem;
        private System.Windows.Forms.ToolStripStatusLabel DoNotClearGridOnReloadLabel;
    }
}

