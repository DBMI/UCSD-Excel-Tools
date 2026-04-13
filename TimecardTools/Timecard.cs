using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Policy;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using Workbook = Microsoft.Office.Interop.Excel.Workbook;
using Worksheet = Microsoft.Office.Interop.Excel.Worksheet;

namespace TimecardTools
{
    internal class Timecard
    {
        private const string seeNextRow = "see next row";
        private Range lastMonthCumulativeHours;
        private int lastRow;
        private DateTime newFileDate;
        private Range thisMonthComments;
        private Range thisMonthCumulativeHours;
        private Range thisMonthEstimatedHours;
        private Range thisMonthNewHours;
        private Workbook thisWorkbook;
        private Range topOfNewHours;

        internal Timecard() { }

        private void BuildGlobals(Worksheet worksheet)
        {
            thisWorkbook = (Workbook)worksheet.Parent;
        }

        private Worksheet CopyToNewSheet(Worksheet lastMonthSheet)
        {
            lastMonthSheet.Copy(Type.Missing, thisWorkbook.Sheets[thisWorkbook.Sheets.Count]);   // copy
            string newSheetName = newFileDate.ToString("MMMM yyyy");
            Worksheet thisMonthSheet = thisWorkbook.Sheets[thisWorkbook.Sheets.Count];
            thisMonthSheet.Name = newSheetName;                                                 // rename
            return thisMonthSheet;
        }

        internal void Extend(Worksheet worksheet)
        {
            // Determine some global values.
            BuildGlobals(worksheet);

            if (!SaveNextMonthVersion()) { return; }

            // Find the latest sheet.
            Worksheet lastMonthSheet = Utilities.FindLastWorksheet(thisWorkbook);

            if (lastMonthSheet == null) { return; }

            // Point to its first cell in column "K".
            lastMonthCumulativeHours = (Range)lastMonthSheet.Cells[2, 11];

            // Copy all to a new sheet with next month name.
            Worksheet thisMonthSheet = CopyToNewSheet(lastMonthSheet);

            // Point to ITS first cell in column "I".
            thisMonthEstimatedHours = (Range)thisMonthSheet.Cells[2, 9];

            // Point to ITS first cell in column "J".
            thisMonthNewHours = (Range)thisMonthSheet.Cells[2, 10];

            // Remember this one for the missing hours formula.
            topOfNewHours = (Range)thisMonthSheet.Cells[2, 10];

            // And first cell in column "K".
            thisMonthCumulativeHours = (Range)thisMonthSheet.Cells[2, 11];
            
            // And to the comments column "L".
            thisMonthComments = (Range)thisMonthSheet.Cells[2, 12];

            // Zero out the values in the new sheet's "Actual hours this month" column.
            ZeroOutNewMonthActualHours();

            // Shift all the formulas in the new sheet.
            UpdateHoursFormulas();

            // Shift the # days formula in the new sheet.
            UpdateMissingHoursFormulas();

            // Save revised workbook.
            thisWorkbook.Save();
        }

        private string NewFilename()
        {
            string filename = thisWorkbook.FullName;
            string directory = System.IO.Path.GetDirectoryName(filename);
            string justTheFilename = System.IO.Path.GetFileNameWithoutExtension(filename);

            // Parse year, month from string like "DFMResearchProjects_Kevin_2024_11.xlsx".
            Regex regex = new Regex(@"(?<preamble>\D+)(_|\s)(?<year>\d{4})(_|\s)(?<month>\d{1,2})$");
            Match match = regex.Match(justTheFilename);

            if (match.Success)
            {
                if (int.TryParse(match.Groups["year"].Value, out int year) &&
                    int.TryParse(match.Groups["month"].Value, out int month))
                {
                    DateTime oldFileDate = new DateTime(year, month, 1);
                    newFileDate = oldFileDate.AddMonths(1);

                    string newFilename = System.IO.Path.Combine(
                    directory,
                    match.Groups["preamble"].Value + "_" +
                    newFileDate.ToString("yyyy_MM") + ".xlsx");

                    return newFilename;
                }
            }

            return string.Empty;
        }

        private bool PastLastRow(string cellContents)
        {
            return !cellContents.Contains("=") && !cellContents.Contains("see next row");
        }

        private bool SaveNextMonthVersion()
        {
            bool success = false;

            // Generate new name.
            string newFilename = NewFilename();

            if (!string.IsNullOrEmpty(newFilename))
            {
                thisWorkbook.SaveAs(newFilename);
                success = true;
            }

            return success;
        }

        // Test the Estimated Hours column for the "See Next Row" marker.
        private bool SkipThisRow(int rowOffset)
        {
            try
            {
                string cellContents = thisMonthEstimatedHours.Offset[rowOffset, 0].Value.ToString();
                return cellContents.ToLower().Contains(seeNextRow);
            }
            catch (Microsoft.CSharp.RuntimeBinder.RuntimeBinderException)
            {
                return false;
            }
        }

        private void UpdateHoursFormulas()
        {
            bool done = false;
            string newFormula;
            string cellContents;
            int rowOffset = 0;

            while (!done) 
            {
                if (!SkipThisRow(rowOffset))
                {
                    newFormula = "=IFERROR('" +
                                lastMonthCumulativeHours.Worksheet.Name + "'!" +
                                lastMonthCumulativeHours.Offset[rowOffset, 0].Address + " + '" +
                                thisMonthNewHours.Worksheet.Name + "'!" +
                                thisMonthNewHours.Offset[rowOffset, 0].Address + ", \"\")";
                    thisMonthCumulativeHours.Offset[rowOffset, 0].Formula = newFormula;
                }

                // Bump down to the next row.
                rowOffset++;

                // Check the comments column for a blank.
                try
                {
                    cellContents = thisMonthComments.Offset[rowOffset, 0].Value.ToString();
                    done = string.IsNullOrEmpty(cellContents);
                }
                catch (Microsoft.CSharp.RuntimeBinder.RuntimeBinderException)
                {
                    done = true;
                }
            }

            lastRow = rowOffset;
        }

        private void UpdateMissingHoursFormulas()
        {
            string newFormula;

            // Insert formula to sum all the new hours entries to compute hours worked so far this month.
            Range lastCellInNewHoursColumn = thisMonthNewHours.Offset[lastRow - 1, 0];
            newFormula = "=SUM(" + topOfNewHours.Address + ":" + lastCellInNewHoursColumn.Address + ")";
            thisMonthNewHours.Offset[lastRow, 0].Formula = newFormula;

            // Insert formula to compute the available work hours so far this month.
            newFormula = "= 8* (NETWORKDAYS(DATE(" +
                        newFileDate.Year.ToString() + ", " +
                        newFileDate.Month.ToString() + ", 1), TODAY()) - 1)";
            thisMonthNewHours.Offset[lastRow + 1, 0].Formula = newFormula;

            // Compute uncharged hours so far this month.
            newFormula = "=" + thisMonthNewHours.Offset[lastRow + 1, 0].Address + "-" +
                                        thisMonthNewHours.Offset[lastRow, 0].Address;
            thisMonthNewHours.Offset[lastRow + 2, 0].Formula = newFormula;
        }

        private void ZeroOutNewMonthActualHours()
        {
            bool done = false;
            int rowOffset = 0;
            string cellContents = string.Empty;

            while (!done)
            {
                thisMonthNewHours.Offset[rowOffset, 0].Formula = null;

                if (!SkipThisRow(rowOffset))
                {
                    thisMonthNewHours.Offset[rowOffset, 0].Value = 0;
                }

                // Bump down to the next row.
                rowOffset++;

                // Check the comments column for a blank.
                try
                {
                    cellContents = thisMonthComments.Offset[rowOffset, 0].Value.ToString();
                    done = string.IsNullOrEmpty(cellContents);
                }
                catch (Microsoft.CSharp.RuntimeBinder.RuntimeBinderException)
                {
                    done = true;
                }
            }
        }
    }
}