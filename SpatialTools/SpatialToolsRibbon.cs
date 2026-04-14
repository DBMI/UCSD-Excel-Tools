using Microsoft.Office.Core;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using SpatialTools.Properties;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;

// TODO:  Follow these steps to enable the Ribbon (XML) item:

// 1: Copy the following code block into the ThisAddin, ThisWorkbook, or ThisDocument class.

//  protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
//  {
//      return new SpatialToolsRibbon();
//  }

// 2. Create callback methods in the "Ribbon Callbacks" region of this class to handle user
//    actions, such as clicking a button. Note: if you have exported this Ribbon from the Ribbon designer,
//    move your code from the event handlers to the callback methods and modify the code to work with the
//    Ribbon extensibility (RibbonX) programming model.

// 3. Assign attributes to the control tags in the Ribbon XML file to identify the appropriate callback methods in your code.  

// For more information, see the Ribbon XML documentation in the Visual Studio Tools for Office Help.


namespace SpatialTools
{
    [ComVisible(true)]
    public class SpatialToolsRibbon : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI ribbon;

        public SpatialToolsRibbon()
        {
        }
        /////////////////////////////////
        ///                           ///
        ///          IMAGES           ///
        ///                           ///
        /////////////////////////////////

        /// <summary>
        /// Lets the @c SpatialToolsRibbon.xml point to the image for the @c address button.
        /// </summary>
        /// <param name="control">Reference to the IRibbonControl object.</param>
        /// <returns>Bitmap</returns>
        public Bitmap addressButton_GetImage(IRibbonControl control)
        {
            return Resources.census;
        }

        /// <summary>
        /// Lets the @c SpatialToolsRibbon.xml point to the image for the @c HPI button.
        /// </summary>
        /// <param name="control">Reference to the IRibbonControl object.</param>
        /// <returns>Bitmap</returns>
        public Bitmap hpiButton_GetImage(IRibbonControl control)
        {
            return Resources.hpi;
        }


        /// <summary>
        /// Lets the @c SpatialToolsRibbon.xml point to the image for the @c AddSvi button.
        /// </summary>
        /// <param name="control">Reference to the IRibbonControl object.</param>
        /// <returns>Bitmap</returns>
        public Bitmap sviButton_GetImage(IRibbonControl control)
        {
            return Resources.ENV_EPHT_social;
        }

        /////////////////////////////////
        ///                           ///
        ///         ACTIONS           ///
        ///                           ///
        /////////////////////////////////

        /// <summary>
        /// When @c AddHPI button is pressed, this method instantiates a @c HpiProcessor object & calls its @c Scan method.
        /// </summary>
        /// <param name="control">Reference to the IRibbonControl object.</param>

        public void OnAddHPI(IRibbonControl control)
        {
            HpiProcessor hpiProcessor = new HpiProcessor();
            Excel.Worksheet wksheet = (Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet;
            hpiProcessor.Scan(wksheet);
        }

        /// <summary>
        /// When @c AddSVI button is pressed, this method instantiates a @c SviProcessor object & calls its @c Scan method.
        /// </summary>
        /// <param name="control">Reference to the IRibbonControl object.</param>

        public void OnAddSVI(IRibbonControl control)
        {
            SviProcessor sviProcessor = new SviProcessor();
            Excel.Worksheet wksheet = (Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet;
            sviProcessor.Scan(wksheet);
        }

        /// <summary>
        /// When @c address button is pressed, this method instantiates a @c AddressToCensusTract object & calls its @c ConvertColumn method.
        /// </summary>
        /// <param name="control">Reference to the IRibbonControl object.</param>

        public void OnConvertAddress(IRibbonControl control)
        {
            AddressToCensusTract converter = new AddressToCensusTract();
            Excel.Worksheet wksheet = (Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet;
            converter.Convert(wksheet);
        }

        #region IRibbonExtensibility Members

        public string GetCustomUI(string ribbonID)
        {
            return GetResourceText("SpatialTools.SpatialToolsRibbon.xml");
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
