using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Autodesk.Navisworks.Api.Plugins;
using Nw = Autodesk.Navisworks.Api;
using Tl = Autodesk.Navisworks.Api.Timeliner;
//This class performs the main task of checking for openings and edges
namespace MTechProject
{
    struct SlabResult
    {
        public string elementCategory;
        public DateTime edgestartDate;
        public DateTime openingstartDate;
        public DateTime stairstartDate;
    }

    class Operate
    {
        //Code for walls
        public SlabResult RoofOperations(string elementID)
        {
            List <DateTime> ListWallEdgeStartDate = new List <DateTime>();
            List<DateTime> ListWallOpeningStartDate = new List<DateTime>();
            List<DateTime> ListWallStairStartDate = new List<DateTime>();
            string caTegory, protectionInfo=null, elementInfo=null;
            DateTime StartDate = new DateTime();
            SlabResult roofwallResult = new SlabResult();
                        
            roofwallResult.elementCategory = null;

            //open the active document and get model items                    
            foreach (Nw.ModelItem modelItem in Nw.Application.ActiveDocument.Models.RootItemDescendantsAndSelf)
            {

                Nw.ModelItem modelParent = modelItem.Parent;

                if (modelItem.HasGeometry)
                {
                    //select all the property categories of the parent element
                    Nw.PropertyCategoryCollection propCatCollection = modelParent.PropertyCategories;

                    foreach (Nw.PropertyCategory propCat in propCatCollection)
                    {
                        //check for the element type to be wall by iterating through property categories
                        if (propCat.DisplayName.ToString() == "Element" && null != propCat.Properties.FindPropertyByDisplayName("Category"))
                        {
                            caTegory = propCat.Properties.FindPropertyByDisplayName("Category").Value.ToString();
                            caTegory = caTegory.Substring(caTegory.IndexOf(":") + 1);

                            if (caTegory == "Walls")
                            {
                                roofwallResult.elementCategory = "Walls";
                                Nw.PropertyCategoryCollection innerpropCatCollection = modelParent.PropertyCategories;

                                // for walls found check for the ehared parameter input from revit
                                foreach (Nw.PropertyCategory innerpropCat in innerpropCatCollection)
                                {

                                    if (innerpropCat.DisplayName.ToString() == "Element" 
                                        && null != innerpropCat.Properties.FindPropertyByDisplayName("ElementInfo") 
                                        && null!=innerpropCat.Properties.FindPropertyByDisplayName("ProtectionInfo"))
                                    {
                                        elementInfo = innerpropCat.Properties.FindPropertyByDisplayName("ElementInfo").Value.ToString();
                                        protectionInfo = innerpropCat.Properties.FindPropertyByDisplayName("ProtectionInfo").Value.ToString();
                                        elementInfo = elementInfo.Substring(elementInfo.IndexOf(":") + 1);
                                        protectionInfo = protectionInfo.Substring(protectionInfo.IndexOf(":") + 1);

                                        //code for side protection type based on the shared parameter
                                        if (elementInfo == elementID && protectionInfo == "ExteriorWall")
                                        {
                                            
                                            Nw.PropertyCategoryCollection innermostpropCatCollection = modelParent.PropertyCategories;

                                            //get the start date of the wall elements and save it to a list
                                            foreach (Nw.PropertyCategory innermostpropCat in innermostpropCatCollection)
                                            {
                                                if (innermostpropCat.DisplayName.ToString() == "TimeLiner" && 
                                                    null != innermostpropCat.Properties.FindPropertyByDisplayName("Attached to Task Start (Planned):1"))
                                                {
                                                    StartDate = (DateTime)innermostpropCat.Properties
                                                        .FindPropertyByDisplayName("Attached to Task Start (Planned):1").Value.ToDateTime();
                                                    ListWallEdgeStartDate.Add(StartDate);
                                                }

                                            }
                                        }

                                        //code for opening protection
                                        if (elementInfo == elementID && protectionInfo == "Opening")
                                        {
                                            Nw.PropertyCategoryCollection innermostpropCatCollection = modelParent.PropertyCategories;

                                            //get the start date of the wall elements and save it to a list
                                            foreach (Nw.PropertyCategory innermostpropCat in innermostpropCatCollection)
                                            {
                                                if (innermostpropCat.DisplayName.ToString() == "TimeLiner" &&
                                                    null != innermostpropCat.Properties.FindPropertyByDisplayName("Attached to Task Start (Planned):1"))
                                                {
                                                    StartDate = (DateTime)innermostpropCat.Properties
                                                        .FindPropertyByDisplayName("Attached to Task Start (Planned):1").Value.ToDateTime();
                                                    ListWallOpeningStartDate.Add(StartDate);
                                                }

                                            }

                                        }

                                        //code for stair side protection
                                        if (elementInfo == elementID && protectionInfo == "Stair")
                                        {
                                            Nw.PropertyCategoryCollection innermostpropCatCollection = modelParent.PropertyCategories;

                                            //get the start date of the wall elements and save it to a list
                                            foreach (Nw.PropertyCategory innermostpropCat in innermostpropCatCollection)
                                            {
                                                if (innermostpropCat.DisplayName.ToString() == "TimeLiner" &&
                                                    null != innermostpropCat.Properties.FindPropertyByDisplayName("Attached to Task Start (Planned):1"))
                                                {
                                                    StartDate = (DateTime)innermostpropCat.Properties
                                                        .FindPropertyByDisplayName("Attached to Task Start (Planned):1").Value.ToDateTime();
                                                    ListWallStairStartDate.Add(StartDate);
                                                }

                                            }

                                        }

                                    }

                                }                                
                            }

                        }

                    }
                }

            }
            // find the largest of the start dates and return it
            if (ListWallEdgeStartDate.Count != 0)
            {
                ListWallEdgeStartDate.Sort();
                roofwallResult.edgestartDate = new DateTime();
                roofwallResult.edgestartDate = ListWallEdgeStartDate.Last();
                ListWallEdgeStartDate.Clear();
            }
            if (ListWallOpeningStartDate.Count != 0)
            {
                ListWallOpeningStartDate.Sort();
                roofwallResult.openingstartDate = new DateTime();
                roofwallResult.openingstartDate = ListWallOpeningStartDate.Last();
                ListWallOpeningStartDate.Clear();
            }
            if (ListWallStairStartDate.Count != 0)
            {
                ListWallStairStartDate.Sort();
                roofwallResult.stairstartDate = new DateTime();
                roofwallResult.stairstartDate = ListWallStairStartDate.Last();
                ListWallStairStartDate.Clear();
            }
            return roofwallResult;

        }

        //Code for Railing
        public SlabResult RailingOperations(String elementID)
        {
            
            List<DateTime> ListRailingedgeStartDate = new List<DateTime>();
            List<DateTime> ListRailingOpeningStartDate = new List<DateTime>();
            List<DateTime> ListRailingStairStartDate = new List<DateTime>();
            string caTegory, elementInfo=null, protectionInfo=null;
            DateTime StartDate = new DateTime();
            SlabResult roofrailingResult = new SlabResult();

            roofrailingResult.elementCategory = null;

            //open the active document and get model items 
            foreach (Nw.ModelItem modelItem in Nw.Application.ActiveDocument.Models.RootItemDescendantsAndSelf)
            {

                foreach (Nw.ModelItem modelAncestors in modelItem.Ancestors)
                {

                    if (modelItem.HasGeometry)
                    {
                        //select all the property categories of the parent element
                        Nw.PropertyCategoryCollection propCatCollection = modelAncestors.PropertyCategories;

                        //check for the element type to be railing by iterating through property categories
                        foreach (Nw.PropertyCategory propCat in propCatCollection)
                        {
                            
                            if (propCat.DisplayName.ToString() == "Element" && null != propCat.Properties.FindPropertyByDisplayName("Category"))
                            {

                                caTegory = propCat.Properties.FindPropertyByDisplayName("Category").Value.ToString();
                                caTegory = caTegory.Substring(caTegory.IndexOf(":") + 1);

                                if (caTegory == "Railings")
                                {
                                    
                                    Nw.PropertyCategoryCollection innerpropCatCollection = modelAncestors.PropertyCategories;
                                    roofrailingResult.elementCategory = "Railing";

                                    // for railings found check for the Shared parameter input from revit
                                    foreach (Nw.PropertyCategory innerpropCat in innerpropCatCollection)
                                    {

                                        if (innerpropCat.DisplayName.ToString() == "Element" 
                                            && null != innerpropCat.Properties.FindPropertyByDisplayName("ElementInfo")
                                            && null != innerpropCat.Properties.FindPropertyByDisplayName("ProtectionInfo"))
                                        {
                                            elementInfo = innerpropCat.Properties.FindPropertyByDisplayName("ElementInfo").Value.ToString();
                                            protectionInfo = innerpropCat.Properties.FindPropertyByDisplayName("ProtectionInfo").Value.ToString();
                                            elementInfo = elementInfo.Substring(elementInfo.IndexOf(":") + 1);
                                            protectionInfo = protectionInfo.Substring(protectionInfo.IndexOf(":") + 1);

                                            //code for side protection type based on the shared parameter
                                            if (elementInfo == elementID && protectionInfo == "ExteriorWall")
                                            {
                                                Nw.PropertyCategoryCollection innermostpropCatCollection = modelAncestors.PropertyCategories;
                                                
                                                //get the start date of the railing elements and save it to a list                                                
                                                foreach (Nw.PropertyCategory innermostpropCat in innermostpropCatCollection)                                                
                                                {                                                       
                                                        
                                                    if (innermostpropCat.DisplayName.ToString() == "TimeLiner" &&                                                                                                               
                                                        null != innermostpropCat.Properties.FindPropertyByDisplayName("Contained in Task Start (Planned):1"))                                                        
                                                    {                                                            
                                                        StartDate = (DateTime)innermostpropCat.Properties
                                                            .FindPropertyByDisplayName("Contained in Task Start (Planned):1").Value.ToDateTime();                                                            
                                                        ListRailingedgeStartDate.Add(StartDate);                                                                                                                    
                                                    }                                                      
  
                                                }
                                            }

                                            //code for opening protection
                                            if (elementInfo == elementID && protectionInfo == "Opening")
                                            {                                                
                                                Nw.PropertyCategoryCollection innermostpropCatCollection = modelAncestors.PropertyCategories;

                                                //get the start date of the railing elements and save it to a list                                                
                                                foreach (Nw.PropertyCategory innermostpropCat in innermostpropCatCollection)
                                                {

                                                    if (innermostpropCat.DisplayName.ToString() == "TimeLiner" &&
                                                        null != innermostpropCat.Properties.FindPropertyByDisplayName("Contained in Task Start (Planned):1"))
                                                    {
                                                        StartDate = (DateTime)innermostpropCat.Properties
                                                            .FindPropertyByDisplayName("Contained in Task Start (Planned):1").Value.ToDateTime();
                                                        ListRailingOpeningStartDate.Add(StartDate);
                                                    }

                                                }

                                            }

                                            //code for slab side protection
                                            if (elementInfo == elementID && protectionInfo == "Stair")
                                            {
                                                Nw.PropertyCategoryCollection innermostpropCatCollection = modelAncestors.PropertyCategories;

                                                //get the start date of the railing elements and save it to a list                                                
                                                foreach (Nw.PropertyCategory innermostpropCat in innermostpropCatCollection)
                                                {

                                                    if (innermostpropCat.DisplayName.ToString() == "TimeLiner" &&
                                                        null != innermostpropCat.Properties.FindPropertyByDisplayName("Contained in Task Start (Planned):1"))
                                                    {
                                                        StartDate = (DateTime)innermostpropCat.Properties
                                                            .FindPropertyByDisplayName("Contained in Task Start (Planned):1").Value.ToDateTime();
                                                        ListRailingStairStartDate.Add(StartDate);
                                                    }

                                                }

                                            }

                                        }

                                    }
                                }

                            }

                        }
                    }

                }
            }
            // find the largest of the start dates and return it
            if (ListRailingedgeStartDate.Count != 0)
            {
                ListRailingedgeStartDate.Sort();
                roofrailingResult.edgestartDate = new DateTime();
                roofrailingResult.edgestartDate = ListRailingedgeStartDate.Last();
                ListRailingedgeStartDate.Clear();
            }

            if (ListRailingOpeningStartDate.Count != 0)
            {
                ListRailingOpeningStartDate.Sort();
                roofrailingResult.openingstartDate = new DateTime();
                roofrailingResult.openingstartDate = ListRailingOpeningStartDate.Last();
                ListRailingOpeningStartDate.Clear();
            }

            if (ListRailingStairStartDate.Count != 0)
            {
                ListRailingStairStartDate.Sort();
                roofrailingResult.stairstartDate = new DateTime();
                roofrailingResult.stairstartDate = ListRailingStairStartDate.Last();
                ListRailingStairStartDate.Clear();
            }

            return roofrailingResult;
        }
        
        //code for window
        public DateTime WallOperations(string elementID)
        {
            List<DateTime> ListWindowStartDate = new List<DateTime>();
            string caTegory, elementInfo=null;
            DateTime StartDate = new DateTime();
            DateTime returnStartDate = new DateTime();

            //open the active document and get model items 
            foreach (Nw.ModelItem modelItem in Nw.Application.ActiveDocument.Models.RootItemDescendantsAndSelf)
            {
                //loop through the ancestors of the element to find properties in case of family element
                foreach (Nw.ModelItem modelAncestors in modelItem.Ancestors)
                {

                    if (modelItem.HasGeometry)
                    {
                        //select all the property categories of the parent element
                        Nw.PropertyCategoryCollection propCatCollection = modelAncestors.PropertyCategories;

                        //check for the element type to be windows by iterating through property categories
                        foreach (Nw.PropertyCategory propCat in propCatCollection)
                        {
                            
                            if (propCat.DisplayName.ToString() == "Element" && null != propCat.Properties.FindPropertyByDisplayName("Category"))
                            {

                                caTegory = propCat.Properties.FindPropertyByDisplayName("Category").Value.ToString();
                                caTegory = caTegory.Substring(caTegory.IndexOf(":") + 1);

                                if (caTegory == "Windows")
                                {
                                    Nw.PropertyCategoryCollection innerpropCatCollection = modelAncestors.PropertyCategories;

                                    // for windows found check for the Shared parameter input from revit
                                    foreach (Nw.PropertyCategory innerpropCat in innerpropCatCollection)
                                    {

                                        if (innerpropCat.DisplayName.ToString() == "Element" 
                                            && null != innerpropCat.Properties.FindPropertyByDisplayName("ElementInfo"))
                                        {
                                            elementInfo = innerpropCat.Properties.FindPropertyByDisplayName("ElementInfo").Value.ToString();                                            
                                            elementInfo = elementInfo.Substring(elementInfo.IndexOf(":") + 1);

                                            //code for side protection type based on the shared parameter
                                            if (elementInfo == elementID )
                                            {
                                                
                                                Nw.PropertyCategoryCollection innermostpropCatCollection = modelAncestors.PropertyCategories;

                                                //get the start date of the window elements and save it to a list    
                                                foreach (Nw.PropertyCategory innermostpropCat in innermostpropCatCollection)
                                                {
                                                    if (innermostpropCat.DisplayName.ToString() == "TimeLiner" &&
                                                        null != innermostpropCat.Properties.FindPropertyByDisplayName("Contained in Task Start (Planned):1"))
                                                    {
                                                        StartDate = (DateTime)innermostpropCat.Properties
                                                            .FindPropertyByDisplayName("Contained in Task Start (Planned):1").Value.ToDateTime();
                                                        ListWindowStartDate.Add(StartDate);
                                                    }

                                                }
                                            }
                                        }

                                    }
                                }

                            }

                        }
                    }

                }
            }
            // find the largest of the start dates and return it
            if (ListWindowStartDate.Count != 0)
            {
                ListWindowStartDate.Sort();
                returnStartDate = ListWindowStartDate.Last();
                ListWindowStartDate.Clear();
            }
            return returnStartDate;
        }

    }
}
