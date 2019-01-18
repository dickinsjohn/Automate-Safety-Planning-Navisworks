using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Autodesk.Navisworks.Api.Plugins;

namespace MTechProject
{
    struct ElementData
    {
        public string ActivityName;        
        public DateTime ProtStartDate;
        public DateTime ProtEndDate;
        public DateTime EleStartDate;
        public DateTime EleEndDate;
        public string ElementType;
        public string Level;
        public double Elevation;
        public string Hazard;
        public string Protection;        
    }

    class Result
    {
        //list with duplicate safety issues
        public static List<ElementData> ListofSafetyIssues = new List<ElementData>();
        
        //temporary list to store safety issues without duplicates
        private static List<ElementData> NoDuplicatesListofSafetyIssues = new List<ElementData>();

        //function to remove duplicates and sort
        public static void RemoveDuplicates()
        {
            NoDuplicatesListofSafetyIssues = Result.ListofSafetyIssues.Distinct().ToList();
            ListofSafetyIssues.Clear();
            ListofSafetyIssues.AddRange(NoDuplicatesListofSafetyIssues);

            //sort the list
            ListofSafetyIssues.Sort((s1, s2) => s1.ProtStartDate.CompareTo(s2.ProtStartDate));
            NoDuplicatesListofSafetyIssues.Clear();

            //load the show report user control
            PluginRecord pr = Autodesk.Navisworks.Api.Application.Plugins.FindPlugin("MTechProject.LoadShowReport.SYSR");

            if (pr != null && pr is DockPanePluginRecord && pr.IsEnabled)
            {
                //The plugin might need to be loaded into the system
                if (pr.LoadedPlugin == null)
                {
                    pr.LoadPlugin();
                }
                //The plugin need to be casted into a more specific type
                DockPanePlugin dpp = pr.LoadedPlugin as DockPanePlugin;
                if (dpp != null)
                {
                    dpp.Visible = !dpp.Visible;
                }
            }

        }

        //function to convert the safety issues list to datatable
        public static DataTable ToDataTable()
        {
            DataTable dataTable = new DataTable();

            dataTable.Columns.Add("ACTIVITY NAME", typeof(string));
            dataTable.Columns.Add("PROTECTION START DATE", typeof(string));
            dataTable.Columns.Add("PROTECTION END DATE", typeof(string));
            dataTable.Columns.Add("ELEMENT START DATE", typeof(string));
            dataTable.Columns.Add("ELEMENT END DATE", typeof(string));
            dataTable.Columns.Add("ELEMENT TYPE", typeof(string));
            dataTable.Columns.Add("LEVEL", typeof(string));
            dataTable.Columns.Add("ELEVATION", typeof(string));
            dataTable.Columns.Add("HAZARD", typeof(string));
            dataTable.Columns.Add("PROTECTION", typeof(string));

            string[] eaachrow = new string[10];

            foreach (ElementData item in ListofSafetyIssues)
            {
                if (item.Elevation >= 1.8 ||item.Elevation<=-1.8)
                {
                    if (null != item.ActivityName)
                        eaachrow[0] = item.ActivityName;
                    else
                        eaachrow[0] = null;

                    if ("01-01-0001" != item.ProtStartDate.ToShortDateString())
                        eaachrow[1] = item.ProtStartDate.ToShortDateString();
                    else
                        eaachrow[1] = null;

                    if ("01-01-0001" != item.ProtEndDate.ToShortDateString())
                        eaachrow[2] = item.ProtEndDate.ToShortDateString();
                    else
                        eaachrow[2] = null;

                    if ("01-01-0001" != item.EleStartDate.ToShortDateString())
                        eaachrow[3] = item.EleStartDate.ToShortDateString();
                    else
                        eaachrow[3] = null;

                    if ("01-01-0001" != item.EleEndDate.ToShortDateString())
                        eaachrow[4] = item.EleEndDate.ToShortDateString();
                    else
                        eaachrow[4] = null;

                    if (null != item.ElementType)
                        eaachrow[5] = item.ElementType;
                    else
                        eaachrow[5] = null;

                    if (null != item.Level)
                        eaachrow[6] = item.Level;
                    else
                        eaachrow[6] = null;

                    if (null != item.Elevation.ToString())
                        eaachrow[7] = item.Elevation.ToString();
                    else
                        eaachrow[7] = null;

                    if (null != item.Hazard)
                        eaachrow[8] = item.Hazard;
                    else
                        eaachrow[8] = null;

                    if (null != item.Protection)
                        eaachrow[9] = item.Protection;
                    else
                        eaachrow[9] = null;

                    dataTable.Rows.Add(eaachrow);
                }                
            }
            
            return dataTable;
        }
    
    }
}
