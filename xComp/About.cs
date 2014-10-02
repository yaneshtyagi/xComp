using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace yCompnents.OfficeTools.xComp
{
    public partial class About : Form
    {
        public About()
        {
            InitializeComponent();
        }

        private void About_Load(object sender, EventArgs e)
        {
            lblAppName.Text = "xComp";
            lblAppSubTitle.Text = "Compare excel sheets ... effectively.";
            lblTeam.Text = "Credits:";
            StringBuilder sbTeamNames = new StringBuilder();
            sbTeamNames.Append("");
            sbTeamNames.Append("Concept:\r\nYanesh Tyagi\r\n\r\n");
            sbTeamNames.Append("Design:\r\nYanesh Tyagi\r\n\r\n");
            sbTeamNames.Append("Development:\r\nYanesh Tyagi\r\n");
            sbTeamNames.Append("http://yaneshtyagi.com\r\n\r\n");
            sbTeamNames.Append("©: Licenced under GPL. You are free to use, distrubite and modify as long as you keep a referecne to the original developer.");
            txtTeamNames.Text = sbTeamNames.ToString();

        }
    }
}