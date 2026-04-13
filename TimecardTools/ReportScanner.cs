using Microsoft.CSharp.RuntimeBinder;
using System.Collections.Generic;
using System.Reflection;
using System.Text.RegularExpressions;
using Range = Microsoft.Office.Interop.Excel.Range;
using Worksheet = Microsoft.Office.Interop.Excel.Worksheet;

namespace TimecardTools
{
    internal class ReportScanner
    {
        private Range commentsColumn;
        private Range cumulativeHoursColumn;
        private Range dateStartedColumn;
        private Range deptOwnerColumn;
        private Range estimatedWorkHoursColumn;
        private Range expectedCompletionColumn;
        private Range goalColumn;
        private Range titleColumn;

        private const string commentsColumnNameHint = "Comments";
        private const string cumulativeHoursColumnNameHint = "Cumulative";
        private const string dateStartedColumnNameHint = "Date Started";
        private const string deptOwnerColumnNameHint = "Dept Owner";
        private const string estimatedWorkHoursColumnNameHint = "Estimated Work Effort";
        private const string expectedCompletionColumnNameHint = "Expected Completion";
        private const string goalColumnNameHint = "Goal";
        private const string titleColumnNameHint = "Title";

        // What to enter when we can't parse the date.
        private const string unknownDateString = "2099-12-31";

        private const string commentsRegexPattern = @"(?<date>\d{1,2}/\d{1,2}/\d{2}):\s*(?<comment>[\w\W]+)";
        private Regex commentsRegex = new Regex(commentsRegexPattern);

        internal ReportScanner()
        {
        }

        private bool FindColumnsByName(Worksheet worksheet)
        {
            commentsColumn = Utilities.TopOfNamedColumn(worksheet, commentsColumnNameHint);
            cumulativeHoursColumn = Utilities.TopOfNamedColumn(worksheet, cumulativeHoursColumnNameHint);
            dateStartedColumn = Utilities.TopOfNamedColumn(worksheet, dateStartedColumnNameHint);
            deptOwnerColumn = Utilities.TopOfNamedColumn(worksheet, deptOwnerColumnNameHint);
            estimatedWorkHoursColumn = Utilities.TopOfNamedColumn(worksheet, estimatedWorkHoursColumnNameHint);
            expectedCompletionColumn = Utilities.TopOfNamedColumn(worksheet, expectedCompletionColumnNameHint);
            goalColumn = Utilities.TopOfNamedColumn(worksheet, goalColumnNameHint);
            titleColumn = Utilities.TopOfNamedColumn(worksheet, titleColumnNameHint);

            return commentsColumn != null &&
                   cumulativeHoursColumn != null &&
                   dateStartedColumn != null &&
                   deptOwnerColumn != null &&
                   estimatedWorkHoursColumn != null &&
                   expectedCompletionColumn != null &&
                   goalColumn != null &&
                   titleColumn != null;
        }

        private Project ScanProjectOneRow(Worksheet worksheet, int rowOffset)
        {
            // In case date is empty or can't be parsed (like "On Hold").
            string dateStarted = unknownDateString;

            try
            {
                dateStarted = dateStartedColumn.Value.Offset[rowOffset, 0].Value.ToString();
            }
            catch (RuntimeBinderException) { }

            string deptOwner = deptOwnerColumn.Value.Offset[rowOffset, 0].Value.ToString();
            string estimatedWorkEffortHours = estimatedWorkHoursColumn.Value.Offset[rowOffset, 0].Value.ToString();

            // In case date is empty or can't be parsed (like "On Hold").
            string expectedCompletion = unknownDateString;

            try
            {
                expectedCompletion = expectedCompletionColumn.Value.Offset[rowOffset, 0].Value.ToString();
            }
            catch (RuntimeBinderException) { }

            string goal = string.Empty;

            // OK if this column is empty.
            try
            {
                goal = goalColumn.Value.Offset[rowOffset, 0].Value.ToString();
            }
            catch (RuntimeBinderException) { }

            string title = titleColumn.Value.Offset[rowOffset, 0].Value.ToString();

            Project project = new Project(dateStarted: dateStarted,
                                          deptOwner: deptOwner,
                                          estimatedWorkEffortHours: estimatedWorkEffortHours,
                                          expectedCompletion: expectedCompletion,
                                          goal: goal,
                                          title: title,
                                          row: rowOffset + 1);
            return project;
        }

        internal Dictionary<string, Range> ScanSheetForProjectCumulativeHours(Worksheet worksheet)
        {
            Dictionary<string, Range> previousMonthCumulativeHoursCells = new Dictionary<string, Range>();

            if (FindColumnsByName(worksheet))
            {
                int rowOffset = 0;
                int numConsecutiveFailures = 0;
                int maxNumFailures = 3;

                while (true)
                {
                    rowOffset++;

                    try
                    {
                        Project project = ScanProjectOneRow(worksheet, rowOffset);

                        if (project.IsValid())
                        {
                            // Remember where we found the cumulative hours value for this project title.
                            if (previousMonthCumulativeHoursCells.ContainsKey(project.Title()))
                            {
                                previousMonthCumulativeHoursCells[project.Title()] = cumulativeHoursColumn.Offset[rowOffset, 0];
                            }
                            else
                            {
                                previousMonthCumulativeHoursCells.Add(project.Title(), cumulativeHoursColumn.Offset[rowOffset, 0]);

                            }
                        }
                    }
                    catch (RuntimeBinderException)
                    {
                        numConsecutiveFailures++;

                        if (numConsecutiveFailures >= maxNumFailures)
                        {
                            break;
                        }
                    }
                }
            }

            return previousMonthCumulativeHoursCells;
        }
    }
}
