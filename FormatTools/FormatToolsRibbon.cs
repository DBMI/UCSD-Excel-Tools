using Microsoft.Office.Core;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Runtime.Serialization;
using System.Text;
using FormatTools.Properties;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;

// TODO:  Follow these steps to enable the Ribbon (XML) item:

// 1: Copy the following code block into the ThisAddin, ThisWorkbook, or ThisDocument class.

//  protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
//  {
//      return new FormatToolsRibbon();
//  }

// 2. Create callback methods in the "Ribbon Callbacks" region of this class to handle user
//    actions, such as clicking a button. Note: if you have exported this Ribbon from the Ribbon designer,
//    move your code from the event handlers to the callback methods and modify the code to work with the
//    Ribbon extensibility (RibbonX) programming model.

// 3. Assign attributes to the control tags in the Ribbon XML file to identify the appropriate callback methods in your code.  

// For more information, see the Ribbon XML documentation in the Visual Studio Tools for Office Help.


namespace FormatTools
{
    [ComVisible(true)]
    public class FormatToolsRibbon : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI ribbon;

        public FormatToolsRibbon()
        {
        }
        /////////////////////////////////
        ///                           ///
        ///          IMAGES           ///
        ///                           ///
        /////////////////////////////////


        /// <summary>
        /// Lets the @c DecsExcelRibbon.xml point to the image for the @c CopyFormatting button.
        /// </summary>
        /// <param name="control">Reference to the IRibbonControl object.</param>
        /// <returns>Bitmap</returns>
        public Bitmap copyFormatButton_GetImage(IRibbonControl control)
        {
            return Resources.copy_formatting;
        }

        /// <summary>
        /// Lets the @c DecsExcelRibbon.xml point to the image for the @c CountWords button.
        /// </summary>
        /// <param name="control">Reference to the IRibbonControl object.</param>
        /// <returns>Bitmap</returns>
        public Bitmap countWordsButton_GetImage(IRibbonControl control)
        {
            return Resources.abacus;
        }

        /// <summary>
        /// Lets the @c DecsExcelRibbon.xml point to the image for the @c ConvertDates button.
        /// </summary>
        /// <param name="control">Reference to the IRibbonControl object.</param>
        /// <returns>Bitmap</returns>
        public Bitmap dateConvertButton_GetImage(IRibbonControl control)
        {
            return Resources.calendar_with_gear;
        }

        /// <summary>
        /// Lets the @c DecsExcelRibbon.xml point to the image for the @c DateToText button.
        /// </summary>
        /// <param name="control">Reference to the IRibbonControl object.</param>
        /// <returns>Bitmap</returns>
        public Bitmap dateToTextButton_GetImage(IRibbonControl control)
        {
            return Resources.calendar;
        }

        /// <summary>
        /// Lets the @c DecsExcelRibbon.xml point to the image for the @c FormatResults button.
        /// </summary>
        /// <param name="control">Reference to the IRibbonControl object.</param>
        /// <returns>Bitmap</returns>
        public Bitmap formatButton_GetImage(IRibbonControl control)
        {
            return Resources.paint_roller;
        }

        /// <summary>
        /// Lets the @c DecsExcelRibbon.xml point to the image for the @c Stripe button.
        /// </summary>
        /// <param name="control">Reference to the IRibbonControl object.</param>
        /// <returns>Bitmap</returns>
        public Bitmap stripeButton_GetImage(IRibbonControl control)
        {
            return Resources.spreadsheet;
        }

        /////////////////////////////////
        ///                           ///
        ///         ACTIONS           ///
        ///                           ///
        /////////////////////////////////

        /// <summary>
        /// When @c ConvertDates button is pressed, this method instantiates a @c MumpsDateConverter object & calls its @c ConvertColumn method.
        /// </summary>
        /// <param name="control">Reference to the IRibbonControl object.</param>

        public void OnConvertDates(IRibbonControl control)
        {
            MumpsDateConverter converter = new MumpsDateConverter();
            Excel.Worksheet wksheet = (Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet;
            converter.ConvertColumn(wksheet);
        }

        /// <summary>
        /// When @c CopyFormatting button is pressed, this method instantiates a @c Formatter object & calls its @c CopyFormat method.
        /// </summary>
        /// <param name="control">Reference to the IRibbonControl object.</param>

        public void OnCopyFormat(IRibbonControl control)
        {
            Formatter formatter = new Formatter();
            Excel.Worksheet wksheet = (Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet;
            formatter.CopyFormat(wksheet);
        }

        /// <summary>
        /// When @c CountWords button is pressed, this method instantiates a @c WordCounter object & calls its @c Scan method.
        /// </summary>
        /// <param name="control">Reference to the IRibbonControl object.</param>

        public void OnCountWords(IRibbonControl control)
        {
            WordCounter wordCounter = new WordCounter();
            Excel.Worksheet wksheet = (Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet;
            wordCounter.Scan(wksheet);
        }

        /// <summary>
        /// When @c CopyFormatting button is pressed, this method instantiates a @c DateConverter object & calls its @c ToText method.
        /// </summary>
        /// <param name="control">Reference to the IRibbonControl object.</param>

        public void OnDatesToText(IRibbonControl control)
        {
            DateConverter converter = new DateConverter();
            Excel.Worksheet wksheet = (Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet;
            converter.ToText(wksheet);
        }

        /// <summary>
        /// When @c FormatResults button is pressed, this method instantiates a @c Formatter object & calls its @c Format method.
        /// </summary>
        /// <param name="control">Reference to the IRibbonControl object.</param>

        public void OnFormat(IRibbonControl control)
        {
            Formatter formatter = new Formatter();
            Excel.Worksheet wksheet = (Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet;
            formatter.Format(wksheet);
        }

        /// <summary>
        /// When @c stripe button is pressed, this method instantiates a @c Striper object & calls its @c Run method.
        /// </summary>
        /// <param name="control">Reference to the IRibbonControl object.</param>

        public void OnStripe(IRibbonControl control)
        {
            Striper striper = new Striper();
            Excel.Worksheet wksheet = (Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet;
            striper.Run(wksheet);
        }

        #region IRibbonExtensibility Members

        public string GetCustomUI(string ribbonID)
        {
            return GetResourceText("FormatTools.FormatToolsRibbon.xml");
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
