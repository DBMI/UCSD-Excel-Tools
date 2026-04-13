using System.Collections.Generic;
using System.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Range = Microsoft.Office.Interop.Excel.Range;
using Workbook = Microsoft.Office.Interop.Excel.Workbook;
using Worksheet = Microsoft.Office.Interop.Excel.Worksheet;

namespace TimecardTools
{
    /**
     * @brief Useful tools
     */
    internal class Utilities
    {
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
        /// Finds last worksheet.
        /// </summary>
        /// <param name="workbook">Active Workbook.</param>

        internal static Worksheet FindLastWorksheet(Workbook workbook)
        {
            List<Worksheet> sheets = new List<Worksheet>();

            foreach (Worksheet sheet in workbook.Worksheets)
            {
                sheets.Add(sheet);
            }

            return sheets.LastOrDefault();
        }        
        /// <summary>
        /// Find the top cell in the named column.
        /// </summary>
        /// <param name="sheet">Active Worksheet.</param>
        /// <param name="columnName">Name of desired column.</param>
        /// <returns>Range?</returns>
        internal static Range TopOfNamedColumn(Worksheet sheet, string columnNameHint)
        {
            Range range = (Range)sheet.Cells[1, 1];
            int lastUsedCol = Utilities.FindLastCol(sheet);

            // Search along row 1.
            for (int col_index = 1; col_index <= lastUsedCol; col_index++)
            {
                string columnName = range.Value.ToString();

                // Remove newlines, etc.
                columnName = columnName.Replace("\n", string.Empty);
                columnName = columnName.Replace("\r", string.Empty);
                columnName = columnName.Replace("\t", string.Empty);

                if (columnName.Contains(columnNameHint))
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
