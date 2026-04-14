using Microsoft.Office.Core;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using MergeTools.Properties;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;

// TODO:  Follow these steps to enable the Ribbon (XML) item:

// 1: Copy the following code block into the ThisAddin, ThisWorkbook, or ThisDocument class.

//  protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
//  {
//      return new MergeToolsRibbon();
//  }

// 2. Create callback methods in the "Ribbon Callbacks" region of this class to handle user
//    actions, such as clicking a button. Note: if you have exported this Ribbon from the Ribbon designer,
//    move your code from the event handlers to the callback methods and modify the code to work with the
//    Ribbon extensibility (RibbonX) programming model.

// 3. Assign attributes to the control tags in the Ribbon XML file to identify the appropriate callback methods in your code.  

// For more information, see the Ribbon XML documentation in the Visual Studio Tools for Office Help.


namespace MergeTools
{
    [ComVisible(true)]
    public class MergeToolsRibbon : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI ribbon;

        public MergeToolsRibbon()
        {
        }
        /////////////////////////////////
        ///                           ///
        ///          IMAGES           ///
        ///                           ///
        /////////////////////////////////

        /// <summary>
        /// Lets the @c DecsExcelRibbon.xml point to the image for the @c ExtractText button.
        /// </summary>
        /// <param name="control">Reference to the IRibbonControl object.</param>
        /// <returns>Bitmap</returns>
        public Bitmap extractTextButton_GetImage(IRibbonControl control)
        {
            return Resources.uncorker;
        }

        /// <summary>
        /// Lets the @c DecsExcelRibbon.xml point to the image for the @c MatchTextButton button.
        /// </summary>
        /// <param name="control">Reference to the IRibbonControl object.</param>
        /// <returns>Bitmap</returns>
        public Bitmap matchTextButton_GetImage(IRibbonControl control)
        {
            return Resources.match_pieces;
        }

        /// <summary>
        /// Lets the @c DecsExcelRibbon.xml point to the image for the @c MergeFiles button.
        /// </summary>
        /// <param name="control">Reference to the IRibbonControl object.</param>
        /// <returns>Bitmap</returns>
        public Bitmap mergeFilesButton_GetImage(IRibbonControl control)
        {
            return Resources.merge_files;
        }

        /// <summary>
        /// Lets the @c DecsExcelRibbon.xml point to the image for the @c MergeRows button.
        /// </summary>
        /// <param name="control">Reference to the IRibbonControl object.</param>
        /// <returns>Bitmap</returns>
        public Bitmap mergeRowsButton_GetImage(IRibbonControl control)
        {
            return Resources.combine_rows;
        }

        /////////////////////////////////
        ///                           ///
        ///         ACTIONS           ///
        ///                           ///
        /////////////////////////////////

        /// <summary>
        /// When @c ExtractMessage button is pressed, instantiates a @c MessageUnpeeler object & calls its @c Scan method.
        /// </summary>
        /// <param name="control">Reference to the IRibbonControl object.</param>

        public void OnExtractText(IRibbonControl control)
        {
            TextExtractor extractor = new TextExtractor();
            Excel.Worksheet wksheet = (Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet;
            extractor.Extract(wksheet);
        }

        /// <summary>
        /// When @c matchText button is pressed, instantiates a @c MatchText object & calls its @c Match method.
        /// </summary>
        /// <param name="control">Reference to the IRibbonControl object.</param>

        public void OnMatchText(IRibbonControl control)
        {
            TextMatcher textMatcher = new TextMatcher();
            textMatcher.Match();
        }

        /// <summary>
        /// When @c MergeFiles button is pressed, this method instantiates a @c FileMerger object & calls its Merge method.
        /// </summary>
        /// <param name="control">Reference to the IRibbonControl object.</param>

        public void OnMergeFiles(IRibbonControl control)
        {
            FileMerger merger = new FileMerger();
            Excel.Worksheet wksheet = (Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet;
            merger.Merge(wksheet);
        }

        /// <summary>
        /// When @c MergeRows button is pressed, this method instantiates a @c MergeRowsForm.
        /// </summary>
        /// <param name="control">Reference to the IRibbonControl object.</param>

        public void OnMergeRows(IRibbonControl control)
        {
            MergeRowsForm form = new MergeRowsForm();
            form.Visible = true;
        }

        #region IRibbonExtensibility Members

        public string GetCustomUI(string ribbonID)
        {
            return GetResourceText("MergeTools.MergeToolsRibbon.xml");
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
