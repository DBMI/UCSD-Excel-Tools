using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using Excel = Microsoft.Office.Interop.Excel;

namespace PlotTools
{
    public enum InsertSide
    {
        Left,
        Right
    }

    internal class Utilities
    {

        /// <summary>
        /// Finds all worksheets.
        /// </summary>
        /// <param name="workbook">Active Workbook.</param>
        internal static List<Worksheet> CollectAllWorksheets(Workbook workbook)
        {
            List<Worksheet> sheets = new List<Worksheet>();

            foreach (Worksheet sheet in workbook.Worksheets)
            {
                sheets.Add(sheet);
            }

            return sheets;
        }
        /// <summary>
        /// Finds last column containing anything.
        /// </summary>
        /// <param name="sheet">Active Worksheet.</param>
        /// <returns>int</returns>
        // https://stackoverflow.com/a/22151620/18749636
        internal static int FindLastCol(Worksheet sheet)
        {
            // Detect Last used Columns, including cells that contain formulas that result in blank values
            return sheet.Cells.Find(
                                    "*",
                                    System.Reflection.Missing.Value,
                                    System.Reflection.Missing.Value,
                                    System.Reflection.Missing.Value,
                                    Excel.XlSearchOrder.xlByColumns,
                                    Excel.XlSearchDirection.xlPrevious,
                                    false,
                                    System.Reflection.Missing.Value,
                                    System.Reflection.Missing.Value).Column;
        }

        /// <summary>
        /// Finds last row containing anything.
        /// </summary>
        /// <param name="sheet">Active Worksheet.</param>
        /// <returns>int</returns>
        internal static int FindLastRow(Worksheet sheet)
        {
            return sheet.Cells.Find(
                                    "*",
                                    System.Reflection.Missing.Value,
                                    Excel.XlFindLookIn.xlValues,
                                    Excel.XlLookAt.xlWhole,
                                    Excel.XlSearchOrder.xlByRows,
                                    Excel.XlSearchDirection.xlPrevious,
                                    false,
                                    System.Reflection.Missing.Value,
                                    System.Reflection.Missing.Value).Row;
        }
        /// <summary>
        /// Pulls all the column names from the first row of a worksheet.
        /// </summary>
        /// <param name="sheet">Active Worksheet.</param>
        /// <returns>List<string></returns>
        internal static List<string> GetColumnNames(Worksheet sheet)
        {
            List<string> names = new List<string>();
            Range range = (Range)sheet.Cells[1, 1];

            int lastUsedCol = Utilities.FindLastCol(sheet);

            // Search along row 1.
            for (int col_index = 1; col_index <= lastUsedCol; col_index++)
            {
                try
                {
                    names.Add(range.Value.ToString());
                }
                catch (Microsoft.CSharp.RuntimeBinder.RuntimeBinderException)
                {
                    break;
                }

                // Move over one column.
                range = range.Offset[0, 1];
            }

            return names;
        }

        /// <summary>
        /// Builds a dictionary linking column names to ranges from the first row of a worksheet.
        /// </summary>
        /// <param name="sheet">Active Worksheet</param>
        /// <param name="namesDesired">List<string></string></param>
        /// <param name="caseSensitive">bool</param>
        /// <returns>Dictionary mapping string -> Range</returns>
        internal static Dictionary<string, Range> GetColumnRangeDictionary(Worksheet sheet,
                                                                           List<string> namesDesired = null,
                                                                           bool caseSensitive = true)
        {
            Dictionary<string, Range> columns = null;

            if (caseSensitive)
            {
                columns = new Dictionary<string, Range>();
            }
            else
            {
                var comparer = StringComparer.OrdinalIgnoreCase;
                columns = new Dictionary<string, Range>(comparer);
            }

            Range range = (Range)sheet.Cells[1, 1];
            int lastUsedCol = FindLastCol(sheet);

            // Search along row 1.
            for (int col_offset = 0; col_offset < lastUsedCol; col_offset++)
            {
                string thisColumnName = string.Empty;

                try
                {
                    thisColumnName = Convert.ToString(range.Offset[0, col_offset].Value);
                    string test = Convert.ToString(range.Offset[0, col_offset].Value2);
                }
                // If there's nothing in this header, move to next column.
                catch (Microsoft.CSharp.RuntimeBinder.RuntimeBinderException)
                {
                    continue;
                }

                // If column name is empty, can't do anything.
                if (string.IsNullOrEmpty(thisColumnName))
                {
                    continue;
                }

                // If namesDesired isn't specified, then we add every column name.
                // Otherwise, see if this column is one of the droids we're looking for.
                if (namesDesired is null || namesDesired.Count == 0 || namesDesired.Contains(thisColumnName))
                {
                    try
                    {
                        columns.Add(thisColumnName, range.Offset[0, col_offset]);
                    }
                    // If there's already a column by this name, skip this one.
                    catch (System.ArgumentException) { }
                }
            }

            return columns;
        }

        /// <summary>
        /// Has the user selected a column? And just one?
        /// </summary>
        /// <param name="application">Excel application</param>
        /// <returns>Range</returns>
        internal static Range GetSelectedCol(Excel.Application application)
        {
            Range rng = (Range)application.Selection;
            Worksheet sheet = application.Selection.Worksheet;
            Range selectedColumn = null;

            // Whole column? Just one? And containing data?
            if (Utilities.HasData(rng.Columns[1]))
            {
                // We want the TOP of the column.
                int columnNumber = rng.Columns[1].Column;
                selectedColumn = (Range)sheet.Cells[1, columnNumber];
            }

            return selectedColumn;
        }

        /// <summary>
        /// Which columns has the user selected?
        /// </summary>
        /// <param name="application">Excel application</param>
        /// <param name="lastRow">Number of last row with data</param>
        /// <returns>List<Range></returns>
        internal static List<Range> GetSelectedCols(Microsoft.Office.Interop.Excel.Application application)
        {
            Range rng = (Range)application.Selection;
            List<Range> selectedColumns = new List<Range>();
            Worksheet sheet = application.Selection.Worksheet;

            foreach (Range col in rng.Columns)
            {
                // Don't add BLANK columns.
                if (Utilities.HasData(col))
                {
                    // Want the TOP of the column.
                    int columnNumber = col.Column;
                    selectedColumns.Add((Range)sheet.Cells[1, columnNumber]);
                }
            }

            return selectedColumns;
        }

        /// <summary>
        /// Does this range have data?
        /// </summary>
        /// <param name="rng">Range to search</param>
        /// <returns>bool</returns>
        internal static bool HasData(Range rng)
        {
            bool hasData = false;
            Range thisCell;
            int rowNumber = 0;

            while (true)
            {
                rowNumber++;
                thisCell = rng.Cells[rowNumber];
                string cell_contents;
                int numConsecutiveFailures = 0;

                try
                {
                    cell_contents = Convert.ToString(thisCell.Value2);
                    hasData = true;
                    break;
                }
                catch (Microsoft.CSharp.RuntimeBinder.RuntimeBinderException)
                {
                    numConsecutiveFailures++;

                    if (numConsecutiveFailures >= 3)
                    {
                        break;
                    }
                }
            }

            return hasData;
        }

        /// <summary>Fills a ListBox object with a list.</summary>
        /// <param name="listBox">ListBox object</param>
        /// <param name="contents">List<string></param>
        internal static void PopulateListBox(System.Windows.Forms.ListBox listBox, List<string> contents, bool enableWhenPopulated = false)
        {
            listBox.Items.Clear();

            foreach (string item in contents)
            {
                listBox.Items.Add(item);
            }

            if (enableWhenPopulated)
            {
                listBox.Enabled = true;
            }
        }

        /// <summary>
        /// Find the top cell in the named column.
        /// </summary>
        /// <param name="sheet">Active Worksheet.</param>
        /// <param name="columnName">Name of desired column.</param>
        /// <returns>Range</returns>
        internal static Range TopOfNamedColumn(Worksheet sheet, string columnName)
        {
            Range range = (Range)sheet.Cells[1, 1];
            int lastUsedCol = Utilities.FindLastCol(sheet);

            // Search along row 1.
            for (int col_index = 1; col_index <= lastUsedCol; col_index++)
            {
                if (range.Value == columnName)
                {
                    return range;
                }

                // Move over one column.
                range = range.Offset[0, 1];
            }

            return null;
        }
    }
}
