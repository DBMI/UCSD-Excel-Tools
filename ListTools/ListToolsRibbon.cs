using Microsoft.Office.Core;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using ListTools.Properties;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;

// TODO:  Follow these steps to enable the Ribbon (XML) item:

// 1: Copy the following code block into the ThisAddin, ThisWorkbook, or ThisDocument class.

//  protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
//  {
//      return new ListToolsRibbon();
//  }

// 2. Create callback methods in the "Ribbon Callbacks" region of this class to handle user
//    actions, such as clicking a button. Note: if you have exported this Ribbon from the Ribbon designer,
//    move your code from the event handlers to the callback methods and modify the code to work with the
//    Ribbon extensibility (RibbonX) programming model.

// 3. Assign attributes to the control tags in the Ribbon XML file to identify the appropriate callback methods in your code.  

// For more information, see the Ribbon XML documentation in the Visual Studio Tools for Office Help.


namespace ListTools
{
    [ComVisible(true)]
    public class ListToolsRibbon : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI ribbon;

        public ListToolsRibbon()
        {
        }
        /////////////////////////////////
        ///                           ///
        ///          IMAGES           ///
        ///                           ///
        /////////////////////////////////

        /// <summary>
        /// Lets the @c DecsExcelRibbon.xml point to the image for the @c ImportList button.
        /// </summary>
        /// <param name="control">Reference to the IRibbonControl object.</param>
        /// <returns>Bitmap</returns>
        public Bitmap buildListButton_GetImage(IRibbonControl control)
        {
            return Resources.clipboard;
        }

        /// <summary>
        /// Lets the @c DecsExcelRibbon.xml point to the image for the @c ChopIntoTabsButton button.
        /// </summary>
        /// <param name="control">Reference to the IRibbonControl object.</param>
        /// <returns>Bitmap</returns>
        public Bitmap chopIntoTabsButton_GetImage(IRibbonControl control)
        {
            return Resources.slice_into_tabs;
        }

        /// <summary>
        /// Lets the @c DecsExcelRibbon.xml point to the image for the @c Histogram button.
        /// </summary>
        /// <param name="control">Reference to the IRibbonControl object.</param>
        /// <returns>Bitmap</returns>
        public Bitmap histogramButton_GetImage(IRibbonControl control)
        {
            return Resources.histogram;
        }

        /// <summary>
        /// Lets the @c DecsExcelRibbon.xml point to the image for the @c lookupNpi button.
        /// </summary>
        /// <param name="control">Reference to the IRibbonControl object.</param>
        /// <returns>Bitmap</returns>
        public Bitmap lookupNpiButton_GetImage(IRibbonControl control)
        {
            return Resources.NPI_Matching;
        }

        /// <summary>
        /// Lets the @c DecsExcelRibbon.xml point to the image for the @c MatchPhysiciansButton button.
        /// </summary>
        /// <param name="control">Reference to the IRibbonControl object.</param>
        /// <returns>Bitmap</returns>
        public Bitmap matchPhysiciansButton_GetImage(IRibbonControl control)
        {
            return Resources.match_people;
        }

        /// <summary>
        /// Lets the @c DecsExcelRibbon.xml point to the image for the @c onCallList button.
        /// </summary>
        /// <param name="control">Reference to the IRibbonControl object.</param>
        /// <returns>Bitmap</returns>
        public Bitmap onCallListButton_GetImage(IRibbonControl control)
        {
            return Resources.on_call;
        }

        /// <summary>
        /// Lets the @c DecsExcelRibbon.xml point to the image for the @c SearchByEmail button.
        /// </summary>
        /// <param name="control">Reference to the IRibbonControl object.</param>
        /// <returns>Bitmap</returns>
        public Bitmap searchByEmailButton_GetImage(IRibbonControl control)
        {
            return Resources.search_by_email;
        }

        /// <summary>
        /// Lets the @c DecsExcelRibbon.xml point to the image for the @c SignalImportButton button.
        /// </summary>
        /// <param name="control">Reference to the IRibbonControl object.</param>
        /// <returns>Bitmap</returns>
        public Bitmap signalImportButton_GetImage(IRibbonControl control)
        {
            return Resources.json;
        }

        /// <summary>
        /// Lets the @c DecsExcelRibbon.xml point to the image for the @c SortTimes button.
        /// </summary>
        /// <param name="control">Reference to the IRibbonControl object.</param>
        /// <returns>Bitmap</returns>
        public Bitmap sortTimesButton_GetImage(IRibbonControl control)
        {
            return Resources.priority;
        }

        /// <summary>
        /// Lets the @c DecsExcelRibbon.xml point to the image for the @c SortTimesSettings button.
        /// </summary>
        /// <param name="control">Reference to the IRibbonControl object.</param>
        /// <returns>Bitmap</returns>
        public Bitmap sortTimesSettingsButton_GetImage(IRibbonControl control)
        {
            return Resources.priority_settings;
        }

        /////////////////////////////////
        ///                           ///
        ///         ACTIONS           ///
        ///                           ///
        /////////////////////////////////

        /// <summary>
        /// When @c BuildList button is pressed, instantiates a @c ListImporter object & calls its @c Scan method.
        /// </summary>
        /// <param name="control">Reference to the IRibbonControl object.</param>

        public void OnBuildList(IRibbonControl control)
        {
            ListImporter importer = new ListImporter();
            Excel.Worksheet wksheet = (Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet;
            importer.Scan(wksheet);
        }

        /// <summary>
        /// When @c ChopList button is pressed, instantiates a @c ListChopper object & calls its @c Scan method.
        /// </summary>
        /// <param name="control">Reference to the IRibbonControl object.</param>

        public void OnChopList(IRibbonControl control)
        {
            ListChopper chopper = new ListChopper();
            Excel.Worksheet wksheet = (Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet;
            chopper.Scan(wksheet);
        }


        /// <summary>
        /// When @c Histogram button is pressed, this method instantiates a @c HistogramBuilder object & calls its @c Build method.
        /// </summary>
        /// <param name="control">Reference to the IRibbonControl object.</param>

        public void OnHistogram(IRibbonControl control)
        {
            HistogramBuilder histogramBuilder = new HistogramBuilder();
            Excel.Worksheet wksheet = (Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet;
            histogramBuilder.Build(wksheet);
        }

        /// <summary>
        /// When @c lookupNpi button is pressed, instantiates a @c NpiLookup object & calls its @c Search method.
        /// </summary>
        /// <param name="control">Reference to the IRibbonControl object.</param>

        public void OnLookupNpi(IRibbonControl control)
        {
            NpiLookup npiLookup = new NpiLookup();
            Excel.Worksheet wksheet = (Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet;
            npiLookup.Search(wksheet);
        }

        /// <summary>
        /// When @c matchPhysicians button is pressed, instantiates a @c MatchPhysicians object & calls its @c Match method.
        /// </summary>
        /// <param name="control">Reference to the IRibbonControl object.</param>

        public void OnMatchPhysicians(IRibbonControl control)
        {
            PhysicianMatcher physicianMatcher = new PhysicianMatcher();
            physicianMatcher.Match();
        }

        /// <summary>
        /// When @c onCallList button is pressed, instantiates a @c OnCallListProcessor object & calls its @c Scan method.
        /// </summary>
        /// <param name="control">Reference to the IRibbonControl object.</param>

        public void OnProcessCallList(IRibbonControl control)
        {
            OnCallListProcessor onCallListProcessor = new OnCallListProcessor();
            Excel.Worksheet wksheet = (Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet;
            onCallListProcessor.Scan(wksheet);
        }

        /// <summary>
        /// When @c SearchByEmail button is pressed, this method instantiates a @c EmailSearcher object & calls its @c Search method.
        /// </summary>
        /// <param name="control">Reference to the IRibbonControl object.</param>

        public void OnSearchByEMail(IRibbonControl control)
        {
            EmailSearcher emailSearcher = new EmailSearcher();
            Excel.Worksheet wksheet = (Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet;
            emailSearcher.Search(wksheet);
        }

        /// <summary>
        /// When @c SignalImport button is pressed, instantiates a @c SignalTimeInNotes object & calls its @c Import method.
        /// </summary>
        /// <param name="control">Reference to the IRibbonControl object.</param>

        public void OnSignalImport(IRibbonControl control)
        {
            ImportSignalData parser = new ImportSignalData();
            parser.Import();
        }

        /// <summary>
        /// When @c SortTimes button is pressed, this method instantiates a @c TimeSorter object & calls its @c Scan method.
        /// </summary>
        /// <param name="control">Reference to the IRibbonControl object.</param>

        public void OnSortTimes(IRibbonControl control)
        {
            TimeSorter timeSorter = new TimeSorter();
            Excel.Worksheet wksheet = (Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet;
            timeSorter.Scan(wksheet);
        }

        /// <summary>
        /// When @c SortTimesSettings button is pressed, this method instantiates a @c TimeSorterSettings object & calls its @c SetThresholds method.
        /// </summary>
        /// <param name="control">Reference to the IRibbonControl object.</param>

        public void OnSortTimesSettings(IRibbonControl control)
        {
            TimeSorterSettings setup = new TimeSorterSettings();
            setup.Set();
        }

        #region IRibbonExtensibility Members

        public string GetCustomUI(string ribbonID)
        {
            return GetResourceText("ListTools.ListToolsRibbon.xml");
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
