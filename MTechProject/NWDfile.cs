using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Autodesk.Navisworks.Api.Plugins;
using Nw = Autodesk.Navisworks.Api;
using Tl = Autodesk.Navisworks.Api.Timeliner;
//This Class opens the NWD file and reads data based on the start date and end date provided. 
//Returns the type of element linked to the activity as string
namespace MTechProject
{
    class NWDfile
    {
        Nw.Document doc = Nw.Application.ActiveDocument;
       
        //Function to open the timeliner document for the active document
        public void DumpTimeLinerInfo()
        {
            //Get Timeliner document of the active document            
            Nw.DocumentParts.IDocumentTimeliner tl = doc.Timeliner;
            Tl.DocumentTimeliner tl_doc = (Tl.DocumentTimeliner)tl;

            //Selecting the timeliner data source from multiple data sources if available
            foreach (Tl.TimelinerDataSource oDS in tl_doc.DataSources)
            {                
                //Loop to display the tasks in timelinerdatasource field. Loop calls a recursive function ShowTask()
                foreach (Tl.TimelinerTask task in tl_doc.Tasks)
                {
                    ShowTasks(task);
                }                                     
            }

            //Clear all the duplicates in the list of safety issues
            Result.RemoveDuplicates();
                        
        }


        TYPESwitch NewTYPESwitch = new TYPESwitch();

        //Function to return the taks and its attached model item to the  TYPESwitch class
        public void ShowTasks(Tl.TimelinerTask Task)
        {
            DateTime PlannedStartDate = new DateTime();
            DateTime PlannedEndDate = new DateTime();
            string activytyDescription = null;
            List<ElementData> TempListofSafetyIssues = new List<ElementData>();
                                                             
            if (Task.Children.Count == 0 )
            {
                PlannedStartDate = (DateTime)Task.PlannedStartDate;
                PlannedEndDate = (DateTime)Task.PlannedEndDate;
                Tl.TimelinerSelection timeLinerSelection = Task.Selection;

                if (timeLinerSelection != null)
                {

                    if (timeLinerSelection.HasExplicitSelection)
                    {
                        activytyDescription = Task.DisplayName;
                        Nw.ModelItemCollection oExplicitSel = timeLinerSelection.ExplicitSelection;
                        Nw.ModelItemEnumerableCollection modelItemEnumerableCollection = oExplicitSel.Descendants;
                        foreach (Nw.ModelItem modelItem in modelItemEnumerableCollection)
                        {
                            NewTYPESwitch.ModelItem = modelItem;
                            NewTYPESwitch.ElementSelector(activytyDescription);
                        }

                    }

                }
                
            }

            // If a task has children/subtasks, call it recursively till the end node is reached
            foreach (Tl.TimelinerTask childTask in Task.Children)
            {
                ShowTasks(childTask);
            }

        }

    }
}
