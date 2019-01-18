using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Autodesk.Navisworks.Api.Plugins;

namespace MTechProject
{
    //Code for Dock pane Plugin
    [PluginAttribute("MTechProject.LoadDockPane", "SYPN", ToolTip = "Click to run the safety check", DisplayName = "Safety Plugin")]
    [DockPanePlugin(365, 220, FixedSize = true, AutoScroll = false)]

    class LoadDockPane: DockPanePlugin
    {
        //Linking User Control Window with the dock pane pluing
        public override Control CreateControlPane()
        {
            SelectionWindow selectionWindow = new SelectionWindow();
            selectionWindow.CreateControl();
            return selectionWindow;
        }

        public override void DestroyControlPane(Control pane)
        {
            pane.Dispose();
        }      
    }
}
