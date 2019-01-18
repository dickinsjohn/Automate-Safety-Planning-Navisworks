using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Autodesk.Navisworks.Api.Plugins;

namespace MTechProject
{
    //Code for Dock pane Plugin
    [PluginAttribute("MTechProject.LoadShowReport", "SYSR", ToolTip = "Show Report", DisplayName = "Report")]
    [DockPanePlugin(1627, 852, FixedSize = true, AutoScroll = false)]

    class LoadShowReport : DockPanePlugin
    {
        //Linking User Control Window with the dock pane pluing
        public override Control CreateControlPane()
        {
            ShowReport ShowReport = new ShowReport();
            ShowReport.CreateControl();
            return ShowReport;
        }

        public override void DestroyControlPane(Control pane)
        {
            pane.Dispose();
        } 
    }
}
