namespace yCompnents.OfficeTools.xComp
{
    partial class About
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
            this.lblAppName = new System.Windows.Forms.Label();
            this.lblAppSubTitle = new System.Windows.Forms.Label();
            this.panel1 = new System.Windows.Forms.Panel();
            this.txtTeamNames = new System.Windows.Forms.TextBox();
            this.lblTeam = new System.Windows.Forms.Label();
            this.btnOK = new System.Windows.Forms.Button();
            this.panel1.SuspendLayout();
            this.SuspendLayout();
            // 
            // lblAppName
            // 
            this.lblAppName.AutoSize = true;
            this.lblAppName.Dock = System.Windows.Forms.DockStyle.Top;
            this.lblAppName.Font = new System.Drawing.Font("Lucida Handwriting", 20.25F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Underline))), System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblAppName.ForeColor = System.Drawing.Color.Maroon;
            this.lblAppName.Location = new System.Drawing.Point(0, 0);
            this.lblAppName.Name = "lblAppName";
            this.lblAppName.Size = new System.Drawing.Size(116, 36);
            this.lblAppName.TabIndex = 0;
            this.lblAppName.Text = "label1";
            // 
            // lblAppSubTitle
            // 
            this.lblAppSubTitle.AutoSize = true;
            this.lblAppSubTitle.Font = new System.Drawing.Font("Calibri", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblAppSubTitle.ForeColor = System.Drawing.Color.Blue;
            this.lblAppSubTitle.Location = new System.Drawing.Point(4, 36);
            this.lblAppSubTitle.Name = "lblAppSubTitle";
            this.lblAppSubTitle.Size = new System.Drawing.Size(50, 19);
            this.lblAppSubTitle.TabIndex = 1;
            this.lblAppSubTitle.Text = "label2";
            // 
            // panel1
            // 
            this.panel1.Controls.Add(this.txtTeamNames);
            this.panel1.Location = new System.Drawing.Point(13, 95);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(267, 138);
            this.panel1.TabIndex = 2;
            // 
            // txtTeamNames
            // 
            this.txtTeamNames.Location = new System.Drawing.Point(4, 4);
            this.txtTeamNames.Multiline = true;
            this.txtTeamNames.Name = "txtTeamNames";
            this.txtTeamNames.ReadOnly = true;
            this.txtTeamNames.ScrollBars = System.Windows.Forms.ScrollBars.Both;
            this.txtTeamNames.Size = new System.Drawing.Size(260, 131);
            this.txtTeamNames.TabIndex = 0;
            // 
            // lblTeam
            // 
            this.lblTeam.AutoSize = true;
            this.lblTeam.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblTeam.Location = new System.Drawing.Point(13, 76);
            this.lblTeam.Name = "lblTeam";
            this.lblTeam.Size = new System.Drawing.Size(41, 13);
            this.lblTeam.TabIndex = 3;
            this.lblTeam.Text = "label1";
            // 
            // btnOK
            // 
            this.btnOK.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.btnOK.Location = new System.Drawing.Point(111, 239);
            this.btnOK.Name = "btnOK";
            this.btnOK.Size = new System.Drawing.Size(75, 23);
            this.btnOK.TabIndex = 4;
            this.btnOK.Text = "OK";
            this.btnOK.UseVisualStyleBackColor = true;
            // 
            // About
            // 
            this.AcceptButton = this.btnOK;
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.CancelButton = this.btnOK;
            this.ClientSize = new System.Drawing.Size(292, 266);
            this.ControlBox = false;
            this.Controls.Add(this.btnOK);
            this.Controls.Add(this.lblTeam);
            this.Controls.Add(this.panel1);
            this.Controls.Add(this.lblAppSubTitle);
            this.Controls.Add(this.lblAppName);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "About";
            this.ShowInTaskbar = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "About xComp";
            this.Load += new System.EventHandler(this.About_Load);
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label lblAppName;
        private System.Windows.Forms.Label lblAppSubTitle;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Label lblTeam;
        private System.Windows.Forms.TextBox txtTeamNames;
        private System.Windows.Forms.Button btnOK;
    }
}