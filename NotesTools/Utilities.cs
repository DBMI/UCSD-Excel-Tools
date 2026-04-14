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

namespace NotesTools
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
        /// Builds a Range object containing the first column of all rows up to the last row containing any data.
        /// </summary>
        /// <param name="sheet">ActiveWorksheet.</param>
        /// <returns>Range</returns>
        internal static Range AllAvailableRows(Worksheet sheet)
        {
            Range firstCell = (Range)sheet.Cells[1, 1];
            Range lastCell = (Range)sheet.Cells[Utilities.FindLastRow(sheet), 1];
            Range allRows = (Range)sheet.Range[firstCell, lastCell];
            return allRows;
        }

        internal static string CleanColumnNamesForSQL(string columnName)
        {
            // Remove stuff that breaks the SQL script.   
            string niceColumnName = columnName.Trim().
                                        Replace(" ", "_").
                                        Replace("/", "_").
                                        Replace("-", "_").
                                        Replace(",", "").
                                        Replace("(", "").
                                        Replace(")", "").
                                        Replace("__", "_").
                                        Replace("__", "_").
                                        Replace("_-_", "_").
                                        ToUpper();
            return niceColumnName;
        }

        /// <summary>
        /// Clears the "Invalid" highlighting & MouseOver eventhandler from a textbox.
        /// </summary>
        /// <param name="textBox">TextBox object</param>

        internal static void ClearRegexInvalid(TextBox textBox)
        {
            if (textBox == null)
                return;

            // Clear any previous highlighting.
            textBox.BackColor = Color.White;

            // Remove the MouseHover EventHandler.
            DetachEvents(textBox);
        }

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
        /// Removes event handlers from a text box.
        /// </summary>
        /// <param name="textBox">Handle to TextBox object</param>        
        public static void DetachEvents(TextBox textBox)
        {
            object objNew = textBox
                .GetType()
                .GetConstructor(new Type[] { })
                .Invoke(new object[] { });
            PropertyInfo propEvents = textBox
                .GetType()
                .GetProperty("Events", BindingFlags.NonPublic | BindingFlags.Instance);

            EventHandlerList eventHandlerList_obj = (EventHandlerList)
                propEvents.GetValue(textBox, null);
            eventHandlerList_obj.Dispose();
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
        /// Tests to see if RegEx pattern has any capture groups.
        /// </summary>
        /// <param name="regexText">Regular Expression</param>
        /// <returns>bool</returns>
        internal static bool HasCaptureGroups(string regexText)
        {
            bool hasCaptureGroups = false;

            // Empty strings are not errors.
            if (!string.IsNullOrEmpty(regexText))
            {
                try
                {
                    Regex regex = new Regex(regexText);

                    // https://learn.microsoft.com/en-us/dotnet/api/system.text.regularexpressions.regex.getgroupnumbers?view=net-8.0
                    int[] groupNumbers = regex.GetGroupNumbers();
                    hasCaptureGroups = groupNumbers.Count() > 1;
                }
                catch (ArgumentException)
                {
                }
            }

            return hasCaptureGroups;
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

        /// <summary>
        /// Tests to see if string is a valid RegEx
        /// </summary>
        /// <param name="regexText">Regular Expression</param>
        /// <returns>RuleValidationResult object</returns>
        internal static RuleValidationResult IsRegexValid(string regexText)
        {
            // Empty strings are not errors.
            if (!string.IsNullOrEmpty(regexText))
            {
                try
                {
                    Regex regex = new Regex(regexText);
                }
                catch (ArgumentException ex)
                {
                    return new RuleValidationResult(ex);
                }
            }

            return new RuleValidationResult();
        }

        /// <summary>
        /// Highlight textbox to show its RegEx is not valid.
        /// </summary>
        /// <param name="textBox">TextBox object.</param>
        /// <param name="message">String used to fill the ToolTip.</param>

        internal static void MarkRegexInvalid(TextBox textBox, string message)
        {
            if (textBox == null)
                return;

            // Highlight box to show RegEx is invalid.
            textBox.BackColor = Color.Pink;

            ToolTip toolTip = FindToolTip(textBox);

            if (toolTip != null)
            {
                Action<object, System.EventArgs> mouseHover = (sender, e) =>
                {
                    toolTip.SetToolTip(textBox, message);
                };

                textBox.MouseHover += new System.EventHandler(mouseHover);
            }
        }

        // How many non-empty strings are present in the list?
        internal static int NumElementsPresent(List<string> values)
        {
            int numNonEmpties = 0;

            foreach (string value in values)
            {
                if (!string.IsNullOrEmpty(value))
                {
                    numNonEmpties++;
                }
            }

            return numNonEmpties;
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
        /// Saves the workbook as revised using a new name.
        /// </summary>
        /// <param name="workbook">Active workbook</param>
        /// <param name="newFilename">Desired new name for file</param>
        /// <param name="justTheFilename">Stub of filename in case we need to synthesize filename with timestamp</param>

        internal static void SaveRevised(Workbook workbook, string newFilename, string justTheFilename)
        {
            try
            {
                workbook.SaveCopyAs(newFilename);
            }
            catch (System.Runtime.InteropServices.COMException)
            {
                newFilename = System.IO.Path.Combine(
                    justTheFilename + "_" + Utilities.GetTimestamp()
                );
                workbook.SaveCopyAs(newFilename);
            }

            MessageBox.Show("Saved in '" + newFilename + "'.");
        }

        // Get the Range defined by these row, column Range objects.
        internal static Range ThisRowThisColumn(Range rowRange, Range columnRange)
        {
            Worksheet sourceSheet = rowRange.Worksheet as Worksheet;
            int rowNumber = rowRange.Row;

            int columnNumber = columnRange.Column;
            Range dataRange = (Range)sourceSheet.Cells[rowNumber, columnNumber];

            return dataRange;
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

        /// <summary>
        /// Creates MessageBox letting user know we didn't find the named column.
        /// </summary>
        /// <param name="columnName">Name of desired column.</param>        
        internal static void WarnColumnNotFound(string columnName)
        {
            string message = "Column '" + columnName + "' not found.";
            string title = "Not Found";
            MessageBoxButtons buttons = MessageBoxButtons.OK;
            DialogResult result = MessageBox.Show(message, title, buttons, MessageBoxIcon.Warning);
        }
    }
}
