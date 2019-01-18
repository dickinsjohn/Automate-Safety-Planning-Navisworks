using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Autodesk.Navisworks.Api.Plugins;
using Nw = Autodesk.Navisworks.Api;
using Tl = Autodesk.Navisworks.Api.Timeliner;

//CODE FOR OPENING TO BE ADDED
//This class performs the soring of elements and then adds the results generated into a list
namespace MTechProject
{
    //structure to store data of each activity temporarily
    
    class TYPESwitch
    {    
        private Nw.ModelItem modelItem;      
        public double elevation;

        //function to set value to ModelItem
        public Nw.ModelItem ModelItem
        {
            get
            {
                return modelItem;
            }
            set
            {
                modelItem = value;
            }
        }

        
        //Function with switch case to select the operations on an element
        public void ElementSelector(string activityDescription)
        {
            Nw.ModelItem modelParent = modelItem.Parent;
            
            string caTegory;                               
             
            if (modelItem.HasGeometry)
            {
                
                Nw.PropertyCategoryCollection propCatCollection = modelParent.PropertyCategories;
                foreach (Nw.PropertyCategory propCat in propCatCollection)
                {
                    if (propCat.DisplayName.ToString() == "Revit Type" 
                        && null != propCat.Properties.FindPropertyByDisplayName("Category").Value)
                    {                        
                        caTegory = propCat.Properties.FindPropertyByDisplayName("Category").Value.ToString();                        
                        caTegory = caTegory.Substring(caTegory.IndexOf(":") + 1);
                        
                        //if element is a roof
                        if (caTegory == "Roofs")
                        {                            
                            RoofSelector(activityDescription);
                        }
                        
                        //if element is a wall
                         if (caTegory == "Walls")
                        {
                            WallSelector(activityDescription);
                        }
                                                
                        //if element is column
                        else if (caTegory == "Columns")
                        {
                            ColumnSelector(activityDescription);
                        }
                        
                        //if element is beam
                         if (caTegory == "Structural Framing")
                        {
                            BeamSelector(activityDescription);                           
                        }
                             
                        //if element is Foundation
                        else if (caTegory == "Structural Foundations")
                        {
                            FoundationSelector(activityDescription);                            
                        }
                        //if element is Floor
                        else if (caTegory == "Floors")
                        {
                            FloorSelector(activityDescription);                            
                        }
                        
                    }
                }
                
                //for families iterate through the ancestors to get the model item type
                foreach (Nw.ModelItem modelAncestors in modelItem.AncestorsAndSelf)
                {
                    Nw.PropertyCategoryCollection familypropCatCollection = modelAncestors.PropertyCategories;
                    foreach (Nw.PropertyCategory familypropCat in familypropCatCollection)
                    {
                        
                        if (familypropCat.DisplayName.ToString() == "Element" 
                            && null != familypropCat.Properties.FindPropertyByDisplayName("Category"))
                        {

                            caTegory = familypropCat.Properties.FindPropertyByDisplayName("Category").Value.ToString();
                            caTegory = caTegory.Substring(caTegory.IndexOf(":") + 1);

                            //if element is window
                            if (caTegory == "Windows")
                            {
                                WindowSelector(activityDescription);                                
                            }

                            //if element is stairs
                            else if (caTegory == "Stairs")
                            {
                                StairSelector(activityDescription);                              
                            }

                            //if element is railing
                            else if (caTegory == "Railings")
                            {
                                RailingSelector(activityDescription);
                            }
                        }

                    }
                }                   
                
            }            
            return;            
        }


        //FUNCTION FOR ROOF
        public void RoofSelector(string activityDescription)
        {
            Nw.ModelItem modelParent = modelItem.Parent;
            Operate elementOperate = new Operate();
            string elementID;
            List<ElementData> TempListofSafetyIssues = new List<ElementData>();
            ElementData eachElementData = new ElementData();

            DateTime tempDateTime = new DateTime();
            eachElementData.EleStartDate = new DateTime();
            eachElementData.EleEndDate = new DateTime();
            eachElementData.ProtStartDate = new DateTime();
            eachElementData.ProtEndDate = new DateTime();

            SlabResult slabwallResult = new SlabResult();
            SlabResult slabrailingResult = new SlabResult();

            Nw.PropertyCategoryCollection innerpropCatCollection = modelParent.PropertyCategories;
            
            //get element start date and end date
            foreach (Nw.PropertyCategory innerpropCat in innerpropCatCollection)
            {
                if (innerpropCat.DisplayName.ToString() == "TimeLiner"
                    && null != innerpropCat.Properties.FindPropertyByDisplayName("Attached to Task Start (Planned):1")
                    && null != innerpropCat.Properties.FindPropertyByDisplayName("Attached to Task End (Planned):1"))
                {
                    tempDateTime = (DateTime)innerpropCat.Properties
                        .FindPropertyByDisplayName("Attached to Task Start (Planned):1").Value.ToDateTime();
                    eachElementData.ProtStartDate = (DateTime)innerpropCat.Properties
                        .FindPropertyByDisplayName("Attached to Task End (Planned):1").Value.ToDateTime();
                    break;
                }
            }


            //get element level and elevation
            foreach (Nw.PropertyCategory innerpropCat in innerpropCatCollection)
            {
                if (innerpropCat.DisplayName.ToString() == "Base Level"
                    && null != innerpropCat.Properties.FindPropertyByDisplayName("Elevation")
                    && null != innerpropCat.Properties.FindPropertyByDisplayName("Name"))
                {
                    //elevation is returned in feet and not as double
                    elevation = innerpropCat.Properties.FindPropertyByDisplayName("Elevation").Value.ToDoubleLength();
                    elevation = elevation * (double)0.3048;

                    eachElementData.Elevation = elevation;
                    eachElementData.Level = innerpropCat.Properties.FindPropertyByDisplayName("Name").Value.ToString();
                    eachElementData.Level = eachElementData.Level.Substring(eachElementData.Level.IndexOf(":") + 1);
                    break;
                }
            }

            //get start date of edge walls and railings
            foreach (Nw.PropertyCategory innerpropCat in innerpropCatCollection)
            {
                if (innerpropCat.DisplayName.ToString() == "Element ID" 
                    && null != innerpropCat.Properties.FindPropertyByDisplayName("Value"))
                {
                    elementID = innerpropCat.Properties.FindPropertyByDisplayName("Value").Value.ToString();
                    elementID = elementID.Substring(elementID.IndexOf(":") + 1);
                    slabwallResult = elementOperate.RoofOperations(elementID);
                    slabrailingResult = elementOperate.RailingOperations(elementID);
                    break;
                }
            }

            //in case of walls along the boundaries
            if (slabwallResult.edgestartDate.ToShortDateString() != "01-01-0001")
            {
                eachElementData.ProtEndDate = slabwallResult.edgestartDate;
                eachElementData.Hazard = "Unprotected Slab Edges";
                eachElementData.Protection = "Provide edge protection along the slab edges where exterior walls are to be built";
                TempListofSafetyIssues.Add(eachElementData);
            }

            //in case of railings along the boundaries
            if (slabrailingResult.edgestartDate.ToShortDateString() != "01-01-0001")
            {
                eachElementData.ProtEndDate = slabrailingResult.edgestartDate;
                eachElementData.Hazard = "Unprotected Edges of Slab";
                eachElementData.Protection = "Provide edge protection along the edges of the slab where railings are to be built";
                TempListofSafetyIssues.Add(eachElementData);
            }

            //in case of walls along openings
            if (slabwallResult.openingstartDate.ToShortDateString() != "01-01-0001")
            {
                eachElementData.ProtEndDate = slabwallResult.openingstartDate;
                eachElementData.Hazard = "Unprotected Openings in slab";
                eachElementData.Protection = "Provide edge protection along the opening edges where walls are to be built";
                TempListofSafetyIssues.Add(eachElementData);
            }

            //in case of railings along openings
            if (slabrailingResult.openingstartDate.ToShortDateString() != "01-01-0001")
            {
                eachElementData.ProtEndDate = slabrailingResult.openingstartDate;
                eachElementData.Hazard = "Unprotected Openings in slab";
                eachElementData.Protection = "Provide edge protection along the edges of openings where railings are to be built";
                TempListofSafetyIssues.Add(eachElementData);
            }


            eachElementData.EleStartDate = tempDateTime;
            eachElementData.EleEndDate = eachElementData.ProtStartDate;
            eachElementData.ProtStartDate = eachElementData.EleEndDate;
            eachElementData.ProtEndDate = eachElementData.EleEndDate;
            eachElementData.ActivityName = activityDescription;
            eachElementData.ElementType = "Roof";
            eachElementData.Hazard = "Unprotected Edges of Slab and slab openings";
            eachElementData.Protection = "See general guidelines for Construction of slabs at height";
            TempListofSafetyIssues.Add(eachElementData);

            Result.ListofSafetyIssues.AddRange(TempListofSafetyIssues);
        }

        //FUNCTION FOR WALL
        public void WallSelector(string activityDescription)
        {
            Nw.ModelItem modelParent = modelItem.Parent;
            Operate elementOperate = new Operate();
            string elementID;
            List<ElementData> TempListofSafetyIssues = new List<ElementData>();
            ElementData eachElementData = new ElementData();

            DateTime tempDateTime = new DateTime();
            eachElementData.EleStartDate = new DateTime();
            eachElementData.EleEndDate = new DateTime();
            eachElementData.ProtStartDate = new DateTime();
            eachElementData.ProtEndDate = new DateTime();
            
            Nw.PropertyCategoryCollection innerpropCatCollection = modelParent.PropertyCategories;
            foreach (Nw.PropertyCategory innerpropCat in innerpropCatCollection)
            {
                if (innerpropCat.DisplayName.ToString() == "Element ID" 
                    && null != innerpropCat.Properties.FindPropertyByDisplayName("Value"))
                {
                    elementID = innerpropCat.Properties.FindPropertyByDisplayName("Value").Value.ToString();
                    elementID = elementID.Substring(elementID.IndexOf(":") + 1);
                    eachElementData.ProtEndDate = elementOperate.WallOperations(elementID);
                    break;
                }
            }

            //get element start date and end date
            foreach (Nw.PropertyCategory innerpropCat in innerpropCatCollection)
            {
                if (innerpropCat.DisplayName.ToString() == "TimeLiner"
                    && null != innerpropCat.Properties.FindPropertyByDisplayName("Attached to Task Start (Planned):1")
                    && null != innerpropCat.Properties.FindPropertyByDisplayName("Attached to Task End (Planned):1"))
                {
                    tempDateTime = (DateTime)innerpropCat.Properties
                        .FindPropertyByDisplayName("Attached to Task Start (Planned):1").Value.ToDateTime();
                    eachElementData.ProtStartDate = (DateTime)innerpropCat
                        .Properties.FindPropertyByDisplayName("Attached to Task End (Planned):1").Value.ToDateTime();
                    break;
                }
            }

            //get element level
            foreach (Nw.PropertyCategory innerpropCat in innerpropCatCollection)
            {
                if (innerpropCat.DisplayName.ToString() == "Base Constraint"
                    && null != innerpropCat.Properties.FindPropertyByDisplayName("Name"))
                {
                    eachElementData.Level = innerpropCat.Properties.FindPropertyByDisplayName("Name").Value.ToString();
                    eachElementData.Level = eachElementData.Level.Substring(eachElementData.Level.IndexOf(":") + 1);
                    break;
                }
            }

            //get element top elevation
            foreach (Nw.PropertyCategory innerpropCat in innerpropCatCollection)
            {
                if (innerpropCat.DisplayName.ToString() == "Top Constraint"
                    && null != innerpropCat.Properties.FindPropertyByDisplayName("Elevation"))
                {
                    //elevation is returned in feet and not as double
                    elevation = innerpropCat.Properties.FindPropertyByDisplayName("Elevation").Value.ToDoubleLength();
                    elevation = elevation * (double)0.3048;

                    eachElementData.Elevation = elevation;
                    break;
                }
            }

            //in case of walls along the boundaries
            if (eachElementData.ProtEndDate.ToShortDateString() != "01-01-0001")
            {
                eachElementData.Hazard = "Window Openings along slab edges";
                eachElementData.Protection = "Close openings using grills";
                TempListofSafetyIssues.Add(eachElementData);
            }

            eachElementData.EleStartDate = tempDateTime;
            eachElementData.EleEndDate = eachElementData.ProtStartDate;
            eachElementData.ProtStartDate = eachElementData.EleEndDate;
            eachElementData.ProtEndDate = eachElementData.EleEndDate;
            eachElementData.ActivityName = activityDescription;
            eachElementData.ElementType = "Wall";
            eachElementData.Hazard = "Working at height";
            eachElementData.Protection = "See general guidelines for Construction of walls at height";
            TempListofSafetyIssues.Add(eachElementData);

            Result.ListofSafetyIssues.AddRange(TempListofSafetyIssues);
        }

        //FUNCTION FOR COLUMN
        public void ColumnSelector(string activityDescription)
        {
            Nw.ModelItem modelParent = modelItem.Parent;
            List<ElementData> TempListofSafetyIssues = new List<ElementData>();
            ElementData eachElementData = new ElementData();

            eachElementData.EleStartDate = new DateTime();
            eachElementData.EleEndDate = new DateTime();
            eachElementData.ProtStartDate = new DateTime();
            eachElementData.ProtEndDate = new DateTime();

            Nw.PropertyCategoryCollection innerpropCatCollection = modelParent.PropertyCategories;

            eachElementData.ActivityName = activityDescription;
            eachElementData.ElementType = "Column";
            eachElementData.Hazard = "Working at height";
            eachElementData.Protection = "See general guidelines for Construction of Columns at height";

            //get element level and elevation
            foreach (Nw.PropertyCategory innerpropCat in innerpropCatCollection)
            {
                if (innerpropCat.DisplayName.ToString() == "Base Level"
                    && null != innerpropCat.Properties.FindPropertyByDisplayName("Name"))
                {
                    eachElementData.Level = innerpropCat.Properties.FindPropertyByDisplayName("Name").Value.ToString();
                    eachElementData.Level = eachElementData.Level.Substring(eachElementData.Level.IndexOf(":") + 1);
                    break;
                }
            }

            //get element top elevation
            foreach (Nw.PropertyCategory innerpropCat in innerpropCatCollection)
            {
                if (innerpropCat.DisplayName.ToString() == "Top Level"
                    && null != innerpropCat.Properties.FindPropertyByDisplayName("Elevation"))
                {
                    //elevation is returned in feet and not as double
                    elevation = innerpropCat.Properties.FindPropertyByDisplayName("Elevation").Value.ToDoubleLength();
                    elevation = elevation * (double)0.3048;

                    eachElementData.Elevation = elevation;
                    break;
                }
            }

            //get element start date and end date
            foreach (Nw.PropertyCategory innerpropCat in innerpropCatCollection)
            {
                if (innerpropCat.DisplayName.ToString() == "TimeLiner"
                    && null != innerpropCat.Properties.FindPropertyByDisplayName("Contained in Task Start (Planned):1")
                    && null != innerpropCat.Properties.FindPropertyByDisplayName("Contained in Task End (Planned):1"))
                {
                    eachElementData.EleStartDate = (DateTime)innerpropCat.Properties
                        .FindPropertyByDisplayName("Contained in Task Start (Planned):1").Value.ToDateTime();
                    eachElementData.EleEndDate = (DateTime)innerpropCat.Properties
                        .FindPropertyByDisplayName("Contained in Task End (Planned):1").Value.ToDateTime();
                    break;
                }
            }
            eachElementData.ProtStartDate = eachElementData.EleStartDate;
            eachElementData.ProtEndDate = eachElementData.EleEndDate;

            TempListofSafetyIssues.Add(eachElementData);

            Result.ListofSafetyIssues.AddRange(TempListofSafetyIssues);
        }

        //FUNCTION FOR BEAM
        public void BeamSelector(string activityDescription)
        {
            Nw.ModelItem modelParent = modelItem.Parent;
            List<ElementData> TempListofSafetyIssues = new List<ElementData>();
            ElementData eachElementData = new ElementData();

            eachElementData.EleStartDate = new DateTime();
            eachElementData.EleEndDate = new DateTime();
            eachElementData.ProtStartDate = new DateTime();
            eachElementData.ProtEndDate = new DateTime();

            Nw.PropertyCategoryCollection innerpropCatCollection = modelParent.PropertyCategories;

            eachElementData.ActivityName = activityDescription;
            eachElementData.ElementType = "Beam";
            eachElementData.Hazard = "Working at height";
            eachElementData.Protection = "See general guidelines for Construction of beams at height";

            //get element level and elevation
            foreach (Nw.PropertyCategory innerpropCat in innerpropCatCollection)
            {
                if (innerpropCat.DisplayName.ToString() == "Reference Level"
                    && null != innerpropCat.Properties.FindPropertyByDisplayName("Elevation")
                    && null != innerpropCat.Properties.FindPropertyByDisplayName("Name"))
                {
                    //elevation is returned in feet and not as double
                    elevation = innerpropCat.Properties.FindPropertyByDisplayName("Elevation").Value.ToDoubleLength();
                    elevation = elevation * (double)0.3048;

                    eachElementData.Elevation = elevation;
                    eachElementData.Level = innerpropCat.Properties.FindPropertyByDisplayName("Name").Value.ToString();
                    eachElementData.Level = eachElementData.Level.Substring(eachElementData.Level.IndexOf(":") + 1);
                    break;
                }
            }

            //get element start date and end date
            foreach (Nw.PropertyCategory innerpropCat in innerpropCatCollection)
            {
                if (innerpropCat.DisplayName.ToString() == "TimeLiner"
                    && null != innerpropCat.Properties.FindPropertyByDisplayName("Contained in Task Start (Planned):1")
                    && null != innerpropCat.Properties.FindPropertyByDisplayName("Contained in Task End (Planned):1"))
                {
                    eachElementData.EleStartDate = (DateTime)innerpropCat.Properties
                        .FindPropertyByDisplayName("Contained in Task Start (Planned):1").Value.ToDateTime();
                    eachElementData.EleEndDate = (DateTime)innerpropCat.Properties
                        .FindPropertyByDisplayName("Contained in Task End (Planned):1").Value.ToDateTime();
                    break;
                }
            }
            eachElementData.ProtStartDate = eachElementData.EleStartDate;
            eachElementData.ProtEndDate = eachElementData.EleEndDate;

            TempListofSafetyIssues.Add(eachElementData);

            Result.ListofSafetyIssues.AddRange(TempListofSafetyIssues);
        }

        //FUNCTION FOR FOUNDATION
        public void FoundationSelector(string activityDescription)
        {
            Nw.ModelItem modelParent = modelItem.Parent;
            List<ElementData> TempListofSafetyIssues = new List<ElementData>();
            ElementData eachElementData = new ElementData();

            eachElementData.EleStartDate = new DateTime();
            eachElementData.EleEndDate = new DateTime();
            eachElementData.ProtStartDate = new DateTime();
            eachElementData.ProtEndDate = new DateTime();

            Nw.PropertyCategoryCollection innerpropCatCollection = modelParent.PropertyCategories;

            eachElementData.ActivityName = activityDescription;
            eachElementData.ElementType = "Foundation";
            eachElementData.Hazard = "Fall into excavation";
            eachElementData.Protection = "See general guidelines for Construction of foundations at depth";

            //get element level and elevation
            foreach (Nw.PropertyCategory innerpropCat in innerpropCatCollection)
            {
                if (innerpropCat.DisplayName.ToString() == "Level"
                    && null != innerpropCat.Properties.FindPropertyByDisplayName("Elevation")
                    && null != innerpropCat.Properties.FindPropertyByDisplayName("Name"))
                {
                    //elevation is returned in feet and not as double
                    elevation = innerpropCat.Properties.FindPropertyByDisplayName("Elevation").Value.ToDoubleLength();
                    elevation = elevation * (double)0.3048;

                    eachElementData.Elevation = elevation;
                    eachElementData.Level = innerpropCat.Properties.FindPropertyByDisplayName("Name").Value.ToString();
                    eachElementData.Level = eachElementData.Level.Substring(eachElementData.Level.IndexOf(":") + 1);
                    break;
                }
            }

            //get element start date and end date
            foreach (Nw.PropertyCategory innerpropCat in innerpropCatCollection)
            {
                if (innerpropCat.DisplayName.ToString() == "TimeLiner"
                    && null != innerpropCat.Properties.FindPropertyByDisplayName("Contained in Task Start (Planned):1")
                    && null != innerpropCat.Properties.FindPropertyByDisplayName("Contained in Task End (Planned):1"))
                {
                    eachElementData.EleStartDate = (DateTime)innerpropCat.Properties
                        .FindPropertyByDisplayName("Contained in Task Start (Planned):1").Value.ToDateTime();
                    eachElementData.EleEndDate = (DateTime)innerpropCat.Properties
                        .FindPropertyByDisplayName("Contained in Task End (Planned):1").Value.ToDateTime();
                    break;
                }
            }
            eachElementData.ProtStartDate = eachElementData.EleStartDate;
            eachElementData.ProtEndDate = eachElementData.EleEndDate;

            TempListofSafetyIssues.Add(eachElementData);

            Result.ListofSafetyIssues.AddRange(TempListofSafetyIssues);
        }

        //FUNCTION FOR FLOOR
        public void FloorSelector(string activityDescription)
        {
            Nw.ModelItem modelParent = modelItem.Parent;
            List<ElementData> TempListofSafetyIssues = new List<ElementData>();
            ElementData eachElementData = new ElementData();

            eachElementData.EleStartDate = new DateTime();
            eachElementData.EleEndDate = new DateTime();
            eachElementData.ProtStartDate = new DateTime();
            eachElementData.ProtEndDate = new DateTime();

            Nw.PropertyCategoryCollection innerpropCatCollection = modelParent.PropertyCategories;

            eachElementData.ActivityName = activityDescription;
            eachElementData.ElementType = "Floor";
            eachElementData.Hazard = "Working at height";
            eachElementData.Protection = "See general guidelines for Construction of Floors at height";

            //get element level and elevation
            foreach (Nw.PropertyCategory innerpropCat in innerpropCatCollection)
            {
                if (innerpropCat.DisplayName.ToString() == "Level"
                    && null != innerpropCat.Properties.FindPropertyByDisplayName("Elevation")
                    && null != innerpropCat.Properties.FindPropertyByDisplayName("Name"))
                {
                    //elevation is returned in feet and not as double
                    elevation = innerpropCat.Properties.FindPropertyByDisplayName("Elevation").Value.ToDoubleLength();
                    elevation = elevation * (double)0.3048;

                    eachElementData.Elevation = elevation;
                    eachElementData.Level = innerpropCat.Properties.FindPropertyByDisplayName("Name").Value.ToString();
                    eachElementData.Level = eachElementData.Level.Substring(eachElementData.Level.IndexOf(":") + 1);
                    break;
                }
            }

            //get element start date and end date
            foreach (Nw.PropertyCategory innerpropCat in innerpropCatCollection)
            {
                if (innerpropCat.DisplayName.ToString() == "TimeLiner"
                    && null != innerpropCat.Properties.FindPropertyByDisplayName("Contained in Task Start (Planned):1")
                    && null != innerpropCat.Properties.FindPropertyByDisplayName("Contained in Task End (Planned):1"))
                {
                    eachElementData.EleStartDate = (DateTime)innerpropCat.Properties
                        .FindPropertyByDisplayName("Contained in Task Start (Planned):1").Value.ToDateTime();
                    eachElementData.EleEndDate = (DateTime)innerpropCat.Properties
                        .FindPropertyByDisplayName("Contained in Task End (Planned):1").Value.ToDateTime();
                    break;
                }
            }
            eachElementData.ProtStartDate = eachElementData.EleStartDate;
            eachElementData.ProtEndDate = eachElementData.EleEndDate;

            TempListofSafetyIssues.Add(eachElementData);
            
            Result.ListofSafetyIssues.AddRange(TempListofSafetyIssues);
        }

        //FUNCTION FOR WINDOW
        public void WindowSelector(string activityDescription)
        {
            Nw.ModelItem modelParent = modelItem.Parent;
            List<ElementData> TempListofSafetyIssues = new List<ElementData>();
            ElementData eachElementData = new ElementData();
            string caTegory = null;

            eachElementData.EleStartDate = new DateTime();
            eachElementData.EleEndDate = new DateTime();
            eachElementData.ProtStartDate = new DateTime();
            eachElementData.ProtEndDate = new DateTime();


            foreach (Nw.ModelItem modelAncestors in modelItem.AncestorsAndSelf)
            {
                Nw.PropertyCategoryCollection familypropCatCollection = modelAncestors.PropertyCategories;
                foreach (Nw.PropertyCategory familypropCat in familypropCatCollection)
                {

                    if (familypropCat.DisplayName.ToString() == "Element"
                        && null != familypropCat.Properties.FindPropertyByDisplayName("Category"))
                    {

                        caTegory = familypropCat.Properties.FindPropertyByDisplayName("Category").Value.ToString();
                        caTegory = caTegory.Substring(caTegory.IndexOf(":") + 1);

                        //if element is window
                        if (caTegory == "Windows")
                        {
                            eachElementData.ActivityName = activityDescription;
                            eachElementData.ElementType = "Window";
                            eachElementData.Protection = "See general guidelines for working at height";

                            //get element start date and end date
                            foreach (Nw.PropertyCategory innermostpropCat in familypropCatCollection)
                            {
                                if (innermostpropCat.DisplayName.ToString() == "TimeLiner"
                                    && null != innermostpropCat.Properties.FindPropertyByDisplayName("Attached to Task Start (Planned):1")
                                    && null != innermostpropCat.Properties.FindPropertyByDisplayName("Attached to Task End (Planned):1"))
                                {
                                    eachElementData.EleStartDate = (DateTime)innermostpropCat.Properties
                                        .FindPropertyByDisplayName("Attached to Task Start (Planned):1").Value.ToDateTime();
                                    eachElementData.EleEndDate = (DateTime)innermostpropCat.Properties
                                        .FindPropertyByDisplayName("Attached to Task End (Planned):1").Value.ToDateTime();
                                    break;
                                }
                            }
                            eachElementData.ProtStartDate = eachElementData.EleStartDate;
                            eachElementData.ProtEndDate = eachElementData.EleEndDate;

                            foreach (Nw.PropertyCategory innerpropCat in familypropCatCollection)
                            {
                                if (innerpropCat.DisplayName.ToString() == "Element"
                                    && null != innerpropCat.Properties.FindPropertyByDisplayName("ElementInfo"))
                                {

                                    eachElementData.Hazard = "Proximity to opening in exterior wall";

                                    //get element level and elevation
                                    foreach (Nw.PropertyCategory innermostpropCat in familypropCatCollection)
                                    {
                                        if (innermostpropCat.DisplayName.ToString() == "Level"
                                            && null != innermostpropCat.Properties.FindPropertyByDisplayName("Elevation")
                                            && null != innermostpropCat.Properties.FindPropertyByDisplayName("Name"))
                                        {
                                            //elevation is returned in feet and not as double
                                            elevation = innermostpropCat.Properties.FindPropertyByDisplayName("Elevation").Value.ToDoubleLength();

                                            eachElementData.Level = innermostpropCat.Properties.FindPropertyByDisplayName("Name").Value.ToString();
                                            eachElementData.Level = eachElementData.Level.Substring(eachElementData.Level.IndexOf(":") + 1);
                                            break;
                                        }
                                    }

                                    //get element sill height
                                    foreach (Nw.PropertyCategory innermostpropCat in familypropCatCollection)
                                    {
                                        if (innermostpropCat.DisplayName.ToString() == "Element"
                                            && null != innermostpropCat.Properties.FindPropertyByDisplayName("Sill Height"))
                                        {
                                            //elevation is returned in feet and not as double
                                            elevation += innermostpropCat.Properties.FindPropertyByDisplayName("Sill Height").Value.ToDoubleLength();
                                            elevation = elevation * (double)0.3048;
                                            eachElementData.Elevation = elevation;
                                            break;
                                        }
                                    }

                                    TempListofSafetyIssues.Add(eachElementData);

                                }
                                else
                                {
                                    eachElementData.Hazard = "Working at Height";
                                    //get element sill height
                                    foreach (Nw.PropertyCategory innermostpropCat in familypropCatCollection)
                                    {
                                        if (innermostpropCat.DisplayName.ToString() == "Element"
                                            && null != innermostpropCat.Properties.FindPropertyByDisplayName("Sill Height"))
                                        {
                                            //elevation is returned in feet and not as double
                                            elevation = innermostpropCat.Properties.FindPropertyByDisplayName("Sill Height").Value.ToDoubleLength();
                                            elevation = elevation * (double)0.3048;
                                            eachElementData.Elevation = elevation;
                                            break;
                                        }
                                    }

                                    TempListofSafetyIssues.Add(eachElementData);

                                }
                            }
                        }
                    }
                }
            }

            Result.ListofSafetyIssues.AddRange(TempListofSafetyIssues);
        }

        //FUNCTION FOR STAIR
        public void StairSelector(string activityDescription)
        {
            string caTegory;
            List<ElementData> TempListofSafetyIssues = new List<ElementData>();
            ElementData eachElementData = new ElementData();
            Operate elementOperate = new Operate();
            string elementID;

            DateTime tempDateTime = new DateTime();
            eachElementData.EleStartDate = new DateTime();
            eachElementData.EleEndDate = new DateTime();
            eachElementData.ProtStartDate = new DateTime();
            eachElementData.ProtEndDate = new DateTime();

            SlabResult stairwallResult = new SlabResult();
            SlabResult stairrailingResult = new SlabResult();

            foreach (Nw.ModelItem modelAncestors in modelItem.AncestorsAndSelf)
            {
                Nw.PropertyCategoryCollection familypropCatCollection = modelAncestors.PropertyCategories;
                foreach (Nw.PropertyCategory familypropCat in familypropCatCollection)
                {

                    if (familypropCat.DisplayName.ToString() == "Element"
                        && null != familypropCat.Properties.FindPropertyByDisplayName("Category"))
                    {

                        caTegory = familypropCat.Properties.FindPropertyByDisplayName("Category").Value.ToString();
                        caTegory = caTegory.Substring(caTegory.IndexOf(":") + 1);

                        //if element is stair
                        if (caTegory == "Stairs")
                        {
                            //get start date of edge walls and railings
                            foreach (Nw.PropertyCategory innerpropCat in familypropCatCollection)
                            {
                                if (innerpropCat.DisplayName.ToString() == "Element ID"
                                    && null != innerpropCat.Properties.FindPropertyByDisplayName("Value"))
                                {
                                    elementID = innerpropCat.Properties.FindPropertyByDisplayName("Value").Value.ToString();
                                    elementID = elementID.Substring(elementID.IndexOf(":") + 1);
                                    stairwallResult = elementOperate.RoofOperations(elementID);
                                    stairrailingResult = elementOperate.RailingOperations(elementID);
                                    break;
                                }
                            }

                            //get element start date and end date
                            foreach (Nw.PropertyCategory innerpropCat in familypropCatCollection)
                            {
                                if (innerpropCat.DisplayName.ToString() == "TimeLiner"
                                    && null != innerpropCat.Properties.FindPropertyByDisplayName("Contained in Task Start (Planned):1")
                                    && null != innerpropCat.Properties.FindPropertyByDisplayName("Contained in Task End (Planned):1"))
                                {
                                    tempDateTime = (DateTime)innerpropCat.Properties
                                        .FindPropertyByDisplayName("Contained in Task Start (Planned):1").Value.ToDateTime();
                                    eachElementData.ProtStartDate = (DateTime)innerpropCat.Properties
                                        .FindPropertyByDisplayName("Contained in Task End (Planned):1").Value.ToDateTime();
                                    break;
                                }
                            }

                            //get element level and elevation
                            foreach (Nw.PropertyCategory innerpropCat in familypropCatCollection)
                            {

                                if (innerpropCat.DisplayName.ToString() == "Top Level"
                                    && null != innerpropCat.Properties.FindPropertyByDisplayName("Elevation")
                                    && null != innerpropCat.Properties.FindPropertyByDisplayName("Name"))
                                {
                                    //elevation is returned in feet and not as double
                                    elevation = innerpropCat.Properties.FindPropertyByDisplayName("Elevation").Value.ToDoubleLength();
                                    elevation = elevation * (double)0.3048;

                                    eachElementData.Elevation = elevation;
                                    eachElementData.Level = innerpropCat.Properties.FindPropertyByDisplayName("Name").Value.ToString();
                                    eachElementData.Level = eachElementData.Level.Substring(eachElementData.Level.IndexOf(":") + 1);
                                    break;
                                }
                            }

                            //in case of walls along the stair edges
                            if (stairwallResult.stairstartDate.ToShortDateString() != "01-01-0001")
                            {
                                eachElementData.ProtEndDate = stairwallResult.stairstartDate;
                                eachElementData.Hazard = "Unprotected Stairs";
                                eachElementData.Protection = "Provide edge protection along the stair walls are to be built";
                                TempListofSafetyIssues.Add(eachElementData);
                            }

                            //in case of railings along the stair edges
                            if (stairrailingResult.stairstartDate.ToShortDateString() != "01-01-0001")
                            {
                                eachElementData.ProtEndDate = stairrailingResult.stairstartDate;
                                eachElementData.Hazard = "Unprotected Stairs";
                                eachElementData.Protection = "Provide edge protection along the stair where railings are to be built";
                                TempListofSafetyIssues.Add(eachElementData);
                            }
                            eachElementData.EleStartDate = tempDateTime;
                            eachElementData.EleEndDate = eachElementData.ProtStartDate;
                            eachElementData.ProtStartDate = eachElementData.EleEndDate;
                            eachElementData.ProtEndDate = eachElementData.EleEndDate;
                            eachElementData.ActivityName = activityDescription;
                            eachElementData.ElementType = "Stair";
                            eachElementData.Hazard = "Working at height";
                            eachElementData.Protection = "See general guidelines for Construction of stairs";
                            TempListofSafetyIssues.Add(eachElementData);

                            Result.ListofSafetyIssues.AddRange(TempListofSafetyIssues);

                        }
                    }
                }
            }
        }

        //FUNCTION FOR RAILING
        public void RailingSelector(string activityDescription)
        {
            List<ElementData> TempListofSafetyIssues = new List<ElementData>();
            ElementData eachElementData = new ElementData();
            string caTegory=null;
            eachElementData.EleStartDate = new DateTime();
            eachElementData.EleEndDate = new DateTime();
            eachElementData.ProtStartDate = new DateTime();
            eachElementData.ProtEndDate = new DateTime();
            
            foreach (Nw.ModelItem modelAncestors in modelItem.AncestorsAndSelf)
            {
                Nw.PropertyCategoryCollection familypropCatCollection = modelAncestors.PropertyCategories;
                foreach (Nw.PropertyCategory familypropCat in familypropCatCollection)
                {

                    if (familypropCat.DisplayName.ToString() == "Element"
                        && null != familypropCat.Properties.FindPropertyByDisplayName("Category"))
                    {

                        caTegory = familypropCat.Properties.FindPropertyByDisplayName("Category").Value.ToString();
                        caTegory = caTegory.Substring(caTegory.IndexOf(":") + 1);

                        //if element is stair
                        if (caTegory == "Railings")
                        {
                            foreach (Nw.PropertyCategory innerpropCat in familypropCatCollection)
                            {

                                if (innerpropCat.DisplayName.ToString() == "Element" && null != innerpropCat.Properties
                                    .FindPropertyByDisplayName("ElementInfo"))
                                {
                                    eachElementData.ActivityName = activityDescription;
                                    eachElementData.ElementType = "Railing";
                                    eachElementData.Hazard = "Proximity to opening/slab edge";
                                    eachElementData.Protection = "See general guidelines working at height";

                                    //get element level and elevation
                                    foreach (Nw.PropertyCategory innermostpropCat in familypropCatCollection)
                                    {
                                        if (innermostpropCat.DisplayName.ToString() == "Base Level"
                                            && null != innermostpropCat.Properties.FindPropertyByDisplayName("Elevation")
                                            && null != innermostpropCat.Properties.FindPropertyByDisplayName("Name"))
                                        {
                                            //elevation is returned in feet and not as double
                                            elevation = innermostpropCat.Properties.FindPropertyByDisplayName("Elevation").Value.ToDoubleLength();
                                            elevation = elevation * (double)0.3048;
                                            eachElementData.Elevation = elevation;

                                            eachElementData.Level = innermostpropCat.Properties.FindPropertyByDisplayName("Name").Value.ToString();
                                            eachElementData.Level = eachElementData.Level.Substring(eachElementData.Level.IndexOf(":") + 1);
                                            break;
                                        }
                                    }

                                    //get element start date and end date
                                    foreach (Nw.PropertyCategory innermostpropCat in familypropCatCollection)
                                    {
                                        if (innermostpropCat.DisplayName.ToString() == "TimeLiner"
                                            && null != innermostpropCat.Properties.FindPropertyByDisplayName("Contained in Task Start (Planned):1")
                                            && null != innermostpropCat.Properties.FindPropertyByDisplayName("Contained in Task End (Planned):1"))
                                        {
                                            eachElementData.EleStartDate = (DateTime)innermostpropCat.Properties
                                                .FindPropertyByDisplayName("Contained in Task Start (Planned):1").Value.ToDateTime();
                                            eachElementData.EleEndDate = (DateTime)innermostpropCat.Properties
                                                .FindPropertyByDisplayName("Contained in Task End (Planned):1").Value.ToDateTime();
                                            break;
                                        }
                                    }
                                    eachElementData.ProtStartDate = eachElementData.EleStartDate;
                                    eachElementData.ProtEndDate = eachElementData.EleEndDate;

                                    TempListofSafetyIssues.Add(eachElementData);
                                }
                            }                            
                        }
                    }
                }
            }
            
            Result.ListofSafetyIssues.AddRange(TempListofSafetyIssues);
        }
              
    }    
}
