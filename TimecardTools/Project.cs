using System;
using System.Collections.Generic;
using System.Text;

namespace TimecardTools
{
    internal class Project
    {
        private DateTime dateStarted;
        private string deptOwner = string.Empty;
        private double estimatedWorkEffortHours = 0;
        private DateTime expectedCompletion;
        private string goal = string.Empty;
        private int row = 0;
        private string title = string.Empty;

        internal Project(string dateStarted, 
                         string deptOwner, 
                         string estimatedWorkEffortHours, 
                         string expectedCompletion,
                         string goal, 
                         int row,
                         string title)
        {
            if (DateTime.TryParse(dateStarted, out DateTime start))
            {
                this.dateStarted = start;
            }

            this.deptOwner = deptOwner;

            if (Double.TryParse(estimatedWorkEffortHours, out double work))
            {
                this.estimatedWorkEffortHours = work;
            }

            if (DateTime.TryParse(expectedCompletion, out DateTime completion))
            {
                this.expectedCompletion = completion;
            }

            this.goal = goal;
            this.row = row;
            this.title = title;
        }

        internal DateTime DateStarted()
        {
            return dateStarted;
        }

        internal string DeptOwner() {return deptOwner; }
        internal double EstimatedWorkEffortHours() {return estimatedWorkEffortHours;}
        internal DateTime ExpectedCompletion() { return expectedCompletion;}
        internal string Goal() {return goal;}
        internal int Row() {return row;}
        internal string Title() {return title;}

        internal bool IsValid()
        {
            return dateStarted > DateTime.MinValue &&
                   expectedCompletion > DateTime.MinValue &&
                   estimatedWorkEffortHours > 0 &&
                   row > 0;
        }
    }
}
