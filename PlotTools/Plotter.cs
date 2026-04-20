using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools.Outlook;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace PlotTools
{
    /**
     * @brief Main class for @c PlottingTools.
     */
    internal class Plotter
    {
        private Excel.Application application;
        private string xColumnName = string.Empty;
        private string yColumnName = string.Empty;

        internal Plotter()
        {
            application = Globals.ThisAddIn.Application;
        }

        private bool FindSelectedColumns(Worksheet worksheet)
        {
            bool success = false;

            // Ask user to select columns of interest.
            Dictionary<string, Range> columns = Utilities.GetColumnRangeDictionary(worksheet);

            using (ChoosePlotVariables form = new ChoosePlotVariables(columns.Keys.ToList<string>()))
            {
                var result = form.ShowDialog();

                if (result == DialogResult.OK)
                {
                    xColumnName = form.xColumnName;
                    yColumnName = form.yColumnName;
                    success = !string.IsNullOrEmpty(xColumnName) && !string.IsNullOrEmpty(yColumnName);
                }
            }

            return success;
        }

        private Range GetNamedColumn(Dictionary<string, Range> columns, string columnName)
        {
            Range columnRange = null;

            if (columns.ContainsKey(columnName))
            {
                columnRange = columns[columnName].EntireColumn;
            }
            else
            {
                string message = "Unable to find column containing '" + columnName + "'.";
                string title = "Not Found";
                MessageBoxButtons buttons = MessageBoxButtons.OK;
                DialogResult result = MessageBox.Show(message, title, buttons, MessageBoxIcon.Warning);
            }

            return columnRange;
        }

        internal void Plot(Worksheet worksheet)
        {
            if (FindSelectedColumns(worksheet))
            {
                PlotOneSheet(worksheet);
            }
        }

        internal void PlotAllSheets(Worksheet worksheet)
        {
            if (FindSelectedColumns(worksheet))
            {
                Workbook workbook = worksheet.Parent as Workbook;
                List<Worksheet> worksheets = Utilities.CollectAllWorksheets(workbook);

                foreach (Worksheet sheet in worksheets)
                {
                    sheet.Select();
                    PlotOneSheet(sheet);
                }
            }
        }

        private void PlotOneSheet(Worksheet worksheet)
        {
            // Ask user to select columns of interest.
            Dictionary<string, Range> columns = Utilities.GetColumnRangeDictionary(worksheet);
            Range xColumnRange = GetNamedColumn(columns, xColumnName);

            if (xColumnRange is null) { return; }

            Range yColumnRange = GetNamedColumn(columns, yColumnName);

            if (yColumnRange is null) { return; }

            Range combinedRange = application.Union(xColumnRange, yColumnRange);

            var chartObject = worksheet.Shapes.AddChart2(-1, Excel.XlChartType.xlXYScatter, 400, 100, 600, 400);
            var chart = chartObject.Chart;
            chart.SetSourceData(combinedRange);
            chart.HasLegend = false;
        }
    }
}
