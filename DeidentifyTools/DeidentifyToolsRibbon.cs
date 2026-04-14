using Microsoft.Office.Core;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Xml.Serialization;
using DeidentifyTools.Properties;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;

// TODO:  Follow these steps to enable the Ribbon (XML) item:

// 1: Copy the following code block into the ThisAddin, ThisWorkbook, or ThisDocument class.

//  protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
//  {
//      return new DeidentifyToolsRibbon();
//  }

// 2. Create callback methods in the "Ribbon Callbacks" region of this class to handle user
//    actions, such as clicking a button. Note: if you have exported this Ribbon from the Ribbon designer,
//    move your code from the event handlers to the callback methods and modify the code to work with the
//    Ribbon extensibility (RibbonX) programming model.

// 3. Assign attributes to the control tags in the Ribbon XML file to identify the appropriate callback methods in your code.  

// For more information, see the Ribbon XML documentation in the Visual Studio Tools for Office Help.


namespace DeidentifyTools
{
    [ComVisible(true)]
    public class DeidentifyToolsRibbon : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI ribbon;

        public DeidentifyToolsRibbon()
        {
        }
        /////////////////////////////////
        ///                           ///
        ///          IMAGES           ///
        ///                           ///
        /////////////////////////////////

        /// <summary>
        /// Lets the @c DecsExcelRibbon.xml point to the image for the @c Bogus button.
        /// </summary>
        /// <param name="control">Reference to the IRibbonControl object.</param>
        /// <returns>Bitmap</returns>
        public Bitmap bogusButton_GetImage(IRibbonControl control)
        {
            return Resources.fake_person;
        }

        /// <summary>
        /// Lets the @c DecsExcelRibbon.xml point to the image for the @c Scrambler button.
        /// </summary>
        /// <param name="control">Reference to the IRibbonControl object.</param>
        /// <returns>Bitmap</returns>
        public Bitmap scramblerButton_GetImage(IRibbonControl control)
        {
            return Resources.groucho;
        }

        /// <summary>
        /// Lets the @c DecsExcelRibbon.xml point to the image for the @c HideDateTime button.
        /// </summary>
        /// <param name="control">Reference to the IRibbonControl object.</param>
        /// <returns>Bitmap</returns>
        public Bitmap hideDateTimeButton_GetImage(IRibbonControl control)
        {
            return Resources.rubber_clock_small;
        }

        /// <summary>
        /// Lets the @c DecsExcelRibbon.xml point to the image for the @c HideNames button.
        /// </summary>
        /// <param name="control">Reference to the IRibbonControl object.</param>
        /// <returns>Bitmap</returns>
        public Bitmap hideNamesButton_GetImage(IRibbonControl control)
        {
            return Resources.hide_identity;
        }
        /////////////////////////////////
        ///                           ///
        ///         ACTIONS           ///
        ///                           ///
        /////////////////////////////////

        /// <summary>
        /// When @c Bogus button is pressed, instantiates a @c ChooseNumBogusRecords object.
        /// </summary>
        /// <param name="control">Reference to the IRibbonControl object.</param>
        public void OnBogus(IRibbonControl control)
        {
            ChooseNumBogusRecordsForm bogusForm = new ChooseNumBogusRecordsForm();
            bogusForm.Visible = true;
        }

        /// <summary>
        /// When @c hideNames button is pressed, instantiates a @c Deidentifier object & calls its @c HideNames method.
        /// </summary>
        /// <param name="control">Reference to the IRibbonControl object.</param>
        public void OnHideNames(IRibbonControl control)
        {
            Deidentifier deidentifier = new Deidentifier();
            Excel.Worksheet wksheet = (Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet;
            deidentifier.HideNames(wksheet);
        }

        /// <summary>
        /// When @c hideDateTime button is pressed, instantiates a @c Deidentifier object & calls its @c ObscureDateTime method.
        /// </summary>
        /// <param name="control">Reference to the IRibbonControl object.</param>
        public void OnObscureDateTime(IRibbonControl control)
        {
            Deidentifier deidentifier = new Deidentifier();
            Excel.Worksheet wksheet = (Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet;
            deidentifier.ObscureDateTime(wksheet);
        }

        /// <summary>
        /// When @c Scrambler button is pressed, this method instantiates a @c Deidentifier object & calls its @c GenerateHash method.
        /// </summary>
        /// <param name="control">Reference to the IRibbonControl object.</param>
        public void OnScramble(IRibbonControl control)
        {
            Deidentifier deidentifier = new Deidentifier();
            Excel.Worksheet wksheet = (Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet;
            deidentifier.GenerateHash(wksheet);
        }

        #region IRibbonExtensibility Members

        public string GetCustomUI(string ribbonID)
        {
            return GetResourceText("DeidentifyTools.DeidentifyToolsRibbon.xml");
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
