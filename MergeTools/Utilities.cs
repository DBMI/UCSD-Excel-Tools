using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text.RegularExpressions;
using System.Web.UI.WebControls;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using TextBox = System.Windows.Forms.TextBox;
using ToolTip = System.Windows.Forms.ToolTip;
using Workbook = Microsoft.Office.Interop.Excel.Workbook;
using Worksheet = Microsoft.Office.Interop.Excel.Worksheet;

namespace MergeTools
{
    public enum InsertSide
    {
        Left,
        Right
    }
    internal enum TypeOfMatch
    {
        DesiredNameStartsWithMatch,
        Exact,
        LastNameOnly,
        Levenshtein,
        MatchStartsWithDesiredName,
        UserSelected
    }
    internal class NameMatch
    {
        private string bestMatch;
        private TypeOfMatch matchType;
        private double? relativeDistance;

        internal NameMatch(string bestMatch, TypeOfMatch matchType, double? relativeDistance = null)
        {
            this.bestMatch = bestMatch;
            this.matchType = matchType;
            this.relativeDistance = relativeDistance;
        }

        internal string BestMatch()
        {
            return bestMatch;
        }
        internal bool IsMatch()
        {
            return !string.IsNullOrEmpty(bestMatch);
        }
        internal string MatchType()
        {
            string thisMatchType = matchType.ToString();

            if (matchType == TypeOfMatch.Levenshtein && relativeDistance.HasValue)
            {
                thisMatchType += ": " + relativeDistance.Value.ToString("F2");
            }

            return thisMatchType;
        }
        internal double? RelativeDistance()
        {
            return relativeDistance;
        }
    }
    /**
     * @brief Useful tools
     */
    internal class Utilities
    {
        /// <summary>
        /// Convert Excel-formatted date to SQL style.
        /// </summary>
        /// <param name="cellContents">String contents of a particular cell.</param>
        /// <returns>DateTime</returns>
        internal static DateTime? ConvertExcelDate(string cellContents)
        {
            DateTime? convertedContents = null;

            if (!string.IsNullOrEmpty(cellContents))
            {
                try
                {
                    double d = double.Parse(cellContents);
                    convertedContents = DateTime.FromOADate(d);
                }
                catch (FormatException)
                {
                    // Try converting directly to DateTime.
                    if (DateTime.TryParse(cellContents, out DateTime result))
                    {
                        convertedContents = result;
                    }
                }
            }

            return convertedContents;
        }

        /// <summary>
        /// Copy row from source Worksheet to target Worksheet.
        /// </summary>
        /// <param name="sourceSheet">Worksheet</param>
        /// <param name="sourceRowOffset">int</param>
        /// <param name="targetSheet">Worksheet</param>
        /// <param name="targetRowOffset">int</param>

        public static void CopyRow(Worksheet sourceSheet, int sourceRowOffset, Worksheet targetSheet, int targetRowOffset)
        {
            // Convert from offset to row number.
            int sourceRowNumber = sourceRowOffset + 1;
            int targetRowNumber = targetRowOffset + 1;

            Range sourceRange = sourceSheet.Rows[sourceRowNumber + ":" + sourceRowNumber];
            Range targetRange = targetSheet.Rows[targetRowNumber + ":" + targetRowNumber];
            sourceRange.Copy(targetRange);
        }

        /// <summary>
        /// Insert new Worksheet with given name.
        /// </summary>
        /// <param name="newName">string</param>

        public static Worksheet CreateNewNamedSheet(string newName)
        {
            Workbook workbook = Globals.ThisAddIn.Application.ActiveWorkbook;
            Worksheet newSheet = null;

            // Create new sheet at the end.
            if (workbook.Sheets.Count > 0)
            {
                newSheet = workbook.Sheets.Add(After: workbook.Sheets[workbook.Sheets.Count]);
            }
            else
            {
                newSheet = workbook.Sheets.Add();
            }

            SafeNamer safeNamer = new SafeNamer(newSheet);
            return safeNamer.AssignName(newName);
        }

        /// <summary>
        /// Given a list of column names ("Coverage Start Date", "Address Start Date"),
        /// finds the strings that make them different ("Address", "Coverage").
        /// </summary>
        /// <param name="columnNames">List of strings</param>
        /// <param name="ignoredWords">List of strings to ignore, like "Start", "Date"</param>
        /// <returns>List of strings</returns>
        internal static List<string> DistinctElements(List<string> columnNames, List<string> ignoredWords)
        {
            List<string> result = new List<string>();

            // Break up all the column names into a list of words, containing duplicates.
            List<string> pieces = new List<string>();

            foreach (string columnName in columnNames)
            {
                pieces = columnName.Split().ToList();

                foreach (string piece in pieces)
                {
                    if (!ignoredWords.Contains(piece) && !result.Contains(piece))
                    {
                        result.Add(piece);
                    }
                }
            }

            return result;
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
        /// <param name="rng">Range</param>
        /// <returns>int</returns>
        // https://stackoverflow.com/a/22151620/18749636
        internal static int FindLastRow(Range rng)
        {
            Worksheet sheet = rng.Worksheet;
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
        /// Finds the ToolTip linked to a TextBox object.
        /// </summary>
        /// <param name="textBox">Handle to TextBox object</param>
        /// <returns>ToolTip</returns>
        internal static ToolTip FindToolTip(TextBox textBox)
        {
            ToolTip toolTip = null;
            Form form = textBox.FindForm();

            if (form != null)
            {
                // https://stackoverflow.com/a/42113517/18749636
                Type typeForm = form.GetType();
                FieldInfo fieldInfo = typeForm.GetField(
                    "components",
                    BindingFlags.Instance | BindingFlags.NonPublic
                );
                IContainer parent = (IContainer)fieldInfo.GetValue(form);
                List<ToolTip> ToolTipList = parent.Components.OfType<ToolTip>().ToList();

                if (ToolTipList.Count > 0)
                {
                    toolTip = ToolTipList[0];
                }
            }

            return toolTip;
        }

        /// <summary>
        /// Turn the scope-of-work filename into a .sql filename.
        /// </summary>
        /// <param name="filename">Scope of work filename</param>
        /// <param name="filenameAddOn">String we want to append to filename</param>
        /// <param name="filetype">Desired filetype (".sql" by default)</param>
        /// <param name="replaceSpaces">Should we replace spaces with underscores (true by default)</param>
        /// <param name="shortVersion">Bool--just filename.type? (false by default)</param>
        /// <returns>string</returns>
        internal static string FormOutputFilename(
            string filename,
            string filenameAddon = "",
            string filetype = ".sql",
            bool replaceSpaces = true,
            bool shortVersion = false
        )
        {
            string dir = Path.GetDirectoryName(filename);
            string justTheFilename = Path.GetFileNameWithoutExtension(filename) + filenameAddon;

            if (replaceSpaces)
            {
                // Make SQL filename import-friendly by replacing spaces with underscores.
                justTheFilename = justTheFilename.Replace(' ', '_');
            }

            string sqlFilename = Path.Combine(dir, justTheFilename + filetype);

            if (shortVersion)
            {
                sqlFilename = justTheFilename + filetype;
            }

            return sqlFilename;
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
            int lastUsedCol = Utilities.FindLastCol(sheet);

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
        /// Builds a dictionary linking column names to ColumnType enum from the first row of a worksheet.
        /// </summary>
        /// <param name="sheet">Active Worksheet.</param>
        /// <returns>Dictionary mapping string -> ColumnType</returns>
        internal static Dictionary<string, ColumnType> GetColumnTypeDictionary(Worksheet sheet)
        {
            Dictionary<string, ColumnType> columns = new Dictionary<string, ColumnType>();
            Range range = (Range)sheet.Cells[1, 1];
            int lastUsedCol = Utilities.FindLastCol(sheet);

            // Search along row 1.
            for (int col_index = 1; col_index <= lastUsedCol; col_index++)
            {
                try
                {
                    string columnName = range.Value.ToString();
                    ColumnType columnType = ColumnType.Text;

                    if (columnName.Contains("Date") || columnName.Contains("DTTM"))
                    {
                        columnType = ColumnType.Date;
                    }

                    columns.Add(range.Value.ToString(), columnType);
                }
                // If there's nothing in this header, then skip it.
                catch (Microsoft.CSharp.RuntimeBinder.RuntimeBinderException) { }
                // If there's already a column by this name, skip this one.
                catch (System.ArgumentException) { }

                // Move over one column.
                range = range.Offset[0, 1];
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
        /// Current time in yyyyMMddHHmmss format
        /// </summary>
        /// <returns>string</returns>
        // https://stackoverflow.com/q/21219797/18749636
        internal static string GetTimestamp()
        {
            return DateTime.Now.ToString("yyyyMMddHHmmss");
        }

        /// <summary>
        /// Finds all the worksheets.
        /// </summary>
        /// <returns>Dictionary<string, Worksheet></returns>
        internal static Dictionary<string, Worksheet> GetWorksheets()
        {
            Workbook workbook = (Workbook)Globals.ThisAddIn.Application.ActiveWorkbook;
            Dictionary<string, Worksheet> dict = new Dictionary<string, Worksheet>();

            foreach (Worksheet worksheet in workbook.Worksheets)
            {
                dict.Add(worksheet.Name, worksheet);
            }

            return dict;
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

        /// <summary>
        /// Inserts new column next to the provided Range.
        /// </summary>
        /// <param name="range">Range of existing column</param>
        /// <param name="newColumnName">Name of column to be created</param>
        /// <param name="side">Create it to the left or right of Range?</param>
        /// <returns>Range</returns>
        internal static Range InsertNewColumn(Range range, string newColumnName, InsertSide side = InsertSide.Right)
        {
            int columnNumber = range.Column;
            Worksheet sheet = range.Worksheet;
            Range newRange;

            if (side == InsertSide.Left)
            {
                sheet.Columns[columnNumber].EntireColumn.Insert();
                newRange = range.Offset[0, -1];
            }
            else
            {
                sheet.Columns[columnNumber + 1].EntireColumn.Insert();
                newRange = range.Offset[0, 1];
            }

            newRange.Cells[1].Value2 = newColumnName;
            return newRange;
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
