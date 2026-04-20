using Microsoft.Office.Core;
using PlotTools.Properties;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;

// TODO:  Follow these steps to enable the Ribbon (XML) item:

// 1: Copy the following code block into the ThisAddin, ThisWorkbook, or ThisDocument class.

//  protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
//  {
//      return new PlotToolsRibbon();
//  }

// 2. Create callback methods in the "Ribbon Callbacks" region of this class to handle user
//    actions, such as clicking a button. Note: if you have exported this Ribbon from the Ribbon designer,
//    move your code from the event handlers to the callback methods and modify the code to work with the
//    Ribbon extensibility (RibbonX) programming model.

// 3. Assign attributes to the control tags in the Ribbon XML file to identify the appropriate callback methods in your code.  

// For more information, see the Ribbon XML documentation in the Visual Studio Tools for Office Help.


namespace PlotTools
{
    [ComVisible(true)]
    public class PlotToolsRibbon : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI ribbon;

        public PlotToolsRibbon()
        {
        }
        /////////////////////////////////
        ///                           ///
        ///          IMAGES           ///
        ///                           ///
        /////////////////////////////////

        /// <summary>
        /// Lets the @c PlotToolsRibbon.xml point to the image for the @c multiPlot button.
        /// </summary>
        /// <param name="control">Reference to the IRibbonControl object.</param>
        /// <returns>Bitmap</returns>
        public Bitmap multiPlotButton_GetImage(IRibbonControl control)
        {
            return Resources.scatter_plot_robot;
        }

        /// <summary>
        /// Lets the @c PlotToolsRibbon.xml point to the image for the @c onePlot button.
        /// </summary>
        /// <param name="control">Reference to the IRibbonControl object.</param>
        /// <returns>Bitmap</returns>
        public Bitmap onePlotButton_GetImage(IRibbonControl control)
        {
            return Resources.scatter_plot;
        }
        /////////////////////////////////
        ///                           ///
        ///         ACTIONS           ///
        ///                           ///
        /////////////////////////////////

        /// <summary>
        /// When @c multiPlot button is pressed, this method instantiates a @c Plotter object & calls its @c PlotAllSheets method.
        /// </summary>
        /// <param name="control">Reference to the IRibbonControl object.</param>

        public void OnPlotAllSheets(IRibbonControl control)
        {
            Plotter plotter = new Plotter();
            Excel.Worksheet wksheet = (Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet;
            plotter.PlotAllSheets(wksheet);
        }

        /// <summary>
        /// When @c onePlot button is pressed, this method instantiates a @c Plotter object & calls its @c Plot method.
        /// </summary>
        /// <param name="control">Reference to the IRibbonControl object.</param>

        public void OnPlotThisSheet(IRibbonControl control)
        {
            Plotter plotter = new Plotter();
            Excel.Worksheet wksheet = (Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet;
            plotter.Plot(wksheet);
        }


        #region IRibbonExtensibility Members

        public string GetCustomUI(string ribbonID)
        {
            return GetResourceText("PlotTools.PlotToolsRibbon.xml");
        }

        #endregion

        #region Ribbon Callbacks
        //Create callback methods here. For more information about adding callback methods, visit https://go.microsoft.com/fwlink/?LinkID=271226

        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            this.ribbon = ribbonUI;
        }

        #endregion

        #region Helpers

        private static string GetResourceText(string resourceName)
        {
            Assembly asm = Assembly.GetExecutingAssembly();
            string[] resourceNames = asm.GetManifestResourceNames();
            for (int i = 0; i < resourceNames.Length; ++i)
            {
                if (string.Compare(resourceName, resourceNames[i], StringComparison.OrdinalIgnoreCase) == 0)
                {
                    using (StreamReader resourceReader = new StreamReader(asm.GetManifestResourceStream(resourceNames[i])))
                    {
                        if (resourceReader != null)
                        {
                            return resourceReader.ReadToEnd();
                        }
                    }
                }
            }
            return null;
        }

        #endregion
    }
}
