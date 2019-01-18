using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Autodesk.Navisworks.Api.Plugins;

namespace MTechProject
{
    //Code to create a plugin
    [PluginAttribute("MTechProject.Safetyplugin", "SYPN", ToolTip = "Click to run the safety check", DisplayName = "Safety Plugin")]

    public class Safetyplugin : AddInPlugin
    {
        public override int Execute(params string[] parameters)
        {            
            // Starting new code for DockPanePlugin
            // Find the plugin first using its pluginId
            PluginRecord pr = Autodesk.Navisworks.Api.Application.Plugins.FindPlugin("MTechProject.LoadDockPane.SYPN");

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
            //End of new code for DockPanePlugin

            return 0;
        }
    }
}
