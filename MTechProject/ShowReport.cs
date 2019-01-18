using System;
using System.IO;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace MTechProject
{
    public partial class ShowReport : UserControl
    {
        public static DataTable textHazard = new DataTable();
        public static DataTable textPrecaution = new DataTable();

        public ShowReport()
        {
            InitializeComponent();
        }

        private void ShowReport_Load(object sender, EventArgs e)
        {
            //list out the safety issues in the datagridview
            DataTable dataTable = new DataTable();
            
            dataTable = Result.ToDataTable();
            dataGridView1.DataSource = dataTable;


        }


        private void button1_Click(object sender, EventArgs e)
        {
            
            string type=null;
            type = comboBox1.Text;
            dataGridView2.DataSource = null;
            dataGridView3.DataSource = null;
            GetDatatable(type);
            dataGridView2.DataSource = ShowReport.textHazard;
            dataGridView3.DataSource = ShowReport.textPrecaution;
        }

        public void GetDatatable(string type)
        {
            string hazardFile = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
            hazardFile = hazardFile + "\\Hazards.txt";

            string precautionsFile = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
            precautionsFile = precautionsFile + "\\Precautions.txt";

            String sLine = "";

            //CODE FOR HAZARD FILE
            try
            {                      
                System.IO.StreamReader FileStream = new System.IO.StreamReader(hazardFile);
                
                sLine = FileStream.ReadLine();

                //The Split Command splits a string into an array, based on the delimiter
                string[] s = sLine.Split(';');
                if (ShowReport.textHazard.Columns.Count == 0)
                    ShowReport.textHazard.Columns.Add(type, typeof(string));
                else
                    ShowReport.textHazard.Rows.Clear();

                while (sLine != null)
                {
                    if (s[0] == type)
                    {
                        for (int j = 1; j < s.Count(); j++)
                        {
                            ShowReport.textHazard.Rows.Add(s[j]);
                        }
                        break;

                    }
                    sLine = FileStream.ReadLine();
                    if(sLine != null)
                        s = sLine.Split(';');
                                    
                }

                FileStream.Close();

            }
            catch (Exception err)
            {
                //Display any errors in a Message Box.
                System.Windows.Forms.MessageBox.Show("Error:  " + err.Message, "Program Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            //CODE FOR PRECAUTION FILE
            try
            {
                System.IO.StreamReader FileStream = new System.IO.StreamReader(precautionsFile);

                sLine = FileStream.ReadLine();

                //The Split Command splits a string into an array, based on the delimiter
                string[] s = sLine.Split(';');
                if (ShowReport.textPrecaution.Columns.Count == 0)
                    ShowReport.textPrecaution.Columns.Add(type, typeof(string));
                else
                    ShowReport.textPrecaution.Rows.Clear();

                while (sLine != null)
                {
                    if (s[0] == type)
                    {
                        for (int j = 1; j < s.Count(); j++)
                        {
                            ShowReport.textPrecaution.Rows.Add(s[j]);
                        }

                    }
                    sLine = FileStream.ReadLine();
                    if(sLine!=null)
                        s = sLine.Split(';');    
                }

                FileStream.Close();

            }
            catch (Exception err)
            {
                //Display any errors in a Message Box.
                System.Windows.Forms.MessageBox.Show("Error:  " + err.Message, "Program Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}
