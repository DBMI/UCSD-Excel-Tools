using Microsoft.Office.Core;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using NotesTools.Properties;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;

// TODO:  Follow these steps to enable the Ribbon (XML) item:

// 1: Copy the following code block into the ThisAddin, ThisWorkbook, or ThisDocument class.

//  protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
//  {
//      return new NotesToolsRibbon();
//  }

// 2. Create callback methods in the "Ribbon Callbacks" region of this class to handle user
//    actions, such as clicking a button. Note: if you have exported this Ribbon from the Ribbon designer,
//    move your code from the event handlers to the callback methods and modify the code to work with the
//    Ribbon extensibility (RibbonX) programming model.

// 3. Assign attributes to the control tags in the Ribbon XML file to identify the appropriate callback methods in your code.  

// For more information, see the Ribbon XML documentation in the Visual Studio Tools for Office Help.


namespace NotesTools
{
    [ComVisible(true)]
    public class NotesToolsRibbon : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI ribbon;

        public NotesToolsRibbon()
        {
        }
        /////////////////////////////////
        ///                           ///
        ///          IMAGES           ///
        ///                           ///
        /////////////////////////////////

        /// <summary>
        /// Lets the @c DecsExcelRibbon.xml point to the image for the @c ExtractMessage button.
        /// </summary>
        /// <param name="control">Reference to the IRibbonControl object.</param>
        /// <returns>Bitmap</returns>
        public Bitmap extractMessageButton_GetImage(IRibbonControl control)
        {
            return Resources.nesting_dolls;
        }

        /// <summary>
        /// Lets the @c DecsExcelRibbon.xml point to the image for the @c MergeRows button.
        /// </summary>
        /// <param name="control">Reference to the IRibbonControl object.</param>
        /// <returns>Bitmap</returns>
        public Bitmap mergeNotesButton_GetImage(IRibbonControl control)
        {
            return Resources.merge_rows;
        }

        /// <summary>
        /// Lets the @c DecsExcelRibbon.xml point to the image for the @c SetupConfig button.
        /// </summary>
        /// <param name="control">Reference to the IRibbonControl object.</param>
        /// <returns>Bitmap</returns>
        public Bitmap notesConfigButton_GetImage(IRibbonControl control)
        {
            return Resources.regex_setup_icon;
        }

        /// <summary>
        /// Lets the @c DecsExcelRibbon.xml point to the image for the @c SearchNotes button.
        /// </summary>
        /// <param name="control">Reference to the IRibbonControl object.</param>
        /// <returns>Bitmap</returns>
        public Bitmap notesSearchButton_GetImage(IRibbonControl control)
        {
            return Resources.regex_search_icon;
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

        public void OnExtractMessage(IRibbonControl control)
        {
            MessageUnpeeler unpeeler = new MessageUnpeeler();
            Excel.Worksheet wksheet = (Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet;
            unpeeler.Scan(wksheet);
        }

        /// <summary>
        /// When @c MergeNotes button is pressed, this method instantiates a @c MergeNotesForm.
        /// </summary>
        /// <param name="control">Reference to the IRibbonControl object.</param>

        public void OnMergeNotes(IRibbonControl control)
        {
            MergeNotesForm form = new MergeNotesForm();
            form.Visible = true;
        }

        /// <summary>
        /// When @c SetupConfig button is pressed, this method instantiates a @c DefineRulesForm object
        /// for the user to review & edit notes parsing rules.
        /// </summary>
        /// <param name="control">Reference to the IRibbonControl object.</param>

        public void OnSearchConfig(IRibbonControl control)
        {
            Excel.Worksheet wksheet = (Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet;
            NotesParser parser = new NotesParser(
                _worksheet: wksheet,
                withConfigFile: false,
                allRows: false
            );
            DefineRulesForm form = new DefineRulesForm(parser);
            form.Visible = true;
        }

        /// <summary>
        /// When @c SearchNotes button is pressed, this method instantiates a @c NotesParser object & calls its @c Parse method.
        /// </summary>
        /// <param name="control">Reference to the IRibbonControl object.</param>

        public void OnSearchNotes(IRibbonControl control)
        {
            Excel.Worksheet wksheet = (Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet;
            NotesParser parser = new NotesParser(_worksheet: wksheet);
            parser.Parse();
        }
        #region IRibbonExtensibility Members

        public string GetCustomUI(string ribbonID)
        {
            return GetResourceText("NotesTools.NotesToolsRibbon.xml");
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
