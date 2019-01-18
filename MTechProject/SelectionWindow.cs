using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace MTechProject
{
    public partial class SelectionWindow : UserControl
    {
        public SelectionWindow()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {            

            Result.ListofSafetyIssues.Clear();
            //Openging the NWD file and invoking the DumpTimelinerInfo() Function
            NWDfile OpenDoc = new NWDfile();
            OpenDoc.DumpTimeLinerInfo();
        }
        
    }
}
