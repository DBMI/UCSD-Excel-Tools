using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Web.UI.WebControls;
using System.Windows.Forms;
using Application = Microsoft.Office.Interop.Excel.Application;
using Workbook = Microsoft.Office.Interop.Excel.Workbook;
using Worksheet = Microsoft.Office.Interop.Excel.Worksheet;


namespace ListTools
{
    // Classes were generated automatically by copying JSON data onto clipboard,
    // then using Visual Studio tool: Edit/Paste Special/Paste JSON As classes.
    public class Rootobject
    {
        public Userwebmetadata UserWebMetadata { get; set; }
        public Datum[] Data { get; set; }

        public List<NamePlusId> Metrics()
        {
            List<NamePlusId> metrics = new List<NamePlusId>();
            List<string> metricNames = new List<string>();

            foreach (Datum datum in Data)
            {
                // What names have already been added?
                metricNames = metrics.Select(metric => metric.Combo()).ToList();

                if (!metricNames.Contains(datum.MetricPlusId().Combo()))
                {
                    metrics.Add(datum.MetricPlusId());
                }
            }

            // Sort by name.
            return metrics.OrderBy(p => p.name).ToList(); ;
        }
    }

    public class Userwebmetadata
    {
        public string UserFirstName { get; set; }
        public string UserLastName { get; set; }
        public string UserEmail { get; set; }
        public string UserID { get; set; }

        // Don't need it & it's in a weird format.
        //public DateTime InstantUTC { get; set; }
    }

    public class Datum
    {
        [JsonPropertyName("EMP CID")]
        public string EMPCID { get; set; }

        [JsonPropertyName("SER CID")]
        public string SERCID { get; set; }

        [JsonPropertyName("Clinician Name")]
        public string ClinicianName { get; set; }

        [JsonPropertyName("Clinician Type")]
        public string ClinicianType { get; set; }

        [JsonPropertyName("Service Area")]
        public string ServiceArea { get; set; }

        public string Department { get; set; }
        public string Specialty { get; set; }

        [JsonPropertyName("User Type")]
        public string UserType { get; set; }

        [JsonPropertyName("Reporting Period Start Date")]
        [JsonConverter(typeof(CustomDateTimeConverter))]
        public DateTime ReportingPeriodStartDate { get; set; }

        [JsonPropertyName("Reporting Period End Date")]
        [JsonConverter(typeof(CustomDateTimeConverter))]
        public DateTime ReportingPeriodEndDate { get; set; }

        public string Metric { get; set; }
        public float Numerator { get; set; }
        public float Denominator { get; set; }
        public float Value { get; set; }

        [JsonPropertyName("Metric ID")]
        public int MetricID { get; set; }

        // Combine name & ID into an object.
        public NamePlusId MetricPlusId()
        {
            return new NamePlusId(Metric, MetricID);
        }
    }

    // Holds both the NAME and ID of a metric.
    // This lets us distinguish among metrics with the same name but different IDs.
    public class NamePlusId
    {
        public string name;
        public int id;

        public NamePlusId(string name, int id)
        {
            this.name = name;
            this.id = id;
        }

        // Show combined name like "Time in In Basket_101".
        public string Combo()
        {
            return this.name + "_" + this.id.ToString();
        }
    }

    internal class PhysicianTable
    {
        private Worksheet sheet;
        private int row;

        internal PhysicianTable()
        {
            this.sheet = null;
            this.row = 0;
        }

        internal PhysicianTable(Worksheet sheet, int row = 1)
        {
            this.sheet = sheet;
            this.row = row;
        }

        internal void IncrementRow()
        {
            row += 1;
        }

        internal int Row()
        {
            return row;
        }

        internal void IncrementRow(int delta)
        {
            row += delta;
        }

        internal Worksheet Sheet()
        {
            return sheet;
        }

        internal Range TopLeft()
        {
            return sheet.Cells[1, 1];
        }
    }

    internal enum SignalNotesFormat
    {
        OneBigSheet,
        SeparateSheets
    }

    internal class ImportSignalData
    {
        private Application application;
        private SignalNotesFormat format = SignalNotesFormat.OneBigSheet;
        private Dictionary<NamePlusId, string> metricsToSheets;
        private Workbook thisWorkbook;

        // Keep track of the last row we used on each sheet.
        private Dictionary<string, int> lastRowUsed;

        // Keep track of the sheet used for each physician AND what row we last used on that sheet.
        private Dictionary<string, PhysicianTable> tables;

        internal ImportSignalData()
        {
            tables = new Dictionary<string, PhysicianTable>();
            application = Globals.ThisAddIn.Application;
            thisWorkbook = Globals.ThisAddIn.Application.ActiveWorkbook;
        }

        private List<string> AskUserForFiles()
        {
            List<string> jsonFiles = new List<string>();

            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                // Because we don't specify an opening directory,
                // the dialog will open in the last directory used.
                openFileDialog.Filter = "JSON files (*.json)|*.json";
                openFileDialog.RestoreDirectory = true;
                openFileDialog.Multiselect = true;

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    // Get the path of specified file.
                    jsonFiles = openFileDialog.FileNames.ToList();
                }
            }

            return jsonFiles;
        }

        private void BuildSheets(Rootobject obj)
        {
            application.StatusBar = "Sorting imported data.";

            // Sort the data objects by metricName, then physician name, then date.
            List<Datum> sortedData = obj.Data.OrderBy(x => x.Metric).
                                              ThenBy(y => y.ClinicianName).
                                              ThenBy(y => y.ReportingPeriodEndDate).ToList();

            if (format == SignalNotesFormat.OneBigSheet)
            {
                PutAllPhysiciansOnSameSheet(sortedData);
            }
            else
            {
                PutEachPhysicianOnOwnSheet(sortedData);
            }

            application.StatusBar = "Ready";
        }

        private PhysicianTable FindOrCreateTable(Datum datum)
        {
            if (tables.ContainsKey(datum.ClinicianName))
            {
                return tables[datum.ClinicianName];
            }

            Worksheet newSheet = Utilities.CreateNewNamedSheet(datum.ClinicianName);
            PhysicianTable table = new PhysicianTable(newSheet);
            Range r = table.TopLeft();
            r.Value2 = "Clinician Name:";
            r.Font.Bold = true;
            r.Offset[0, 1].Value2 = datum.ClinicianName;

            r.Offset[1, 0].Value2 = "Clinician Type:";
            r.Offset[1, 0].Font.Bold = true;
            r.Offset[1, 1].Value2 = datum.ClinicianType;

            r.Offset[2, 0].Value2 = "Service Area:";
            r.Offset[2, 0].Font.Bold = true;
            r.Offset[2, 1].Value2 = datum.ServiceArea;

            r.Offset[3, 0].Value2 = "Department:";
            r.Offset[3, 0].Font.Bold = true;
            r.Offset[3, 1].Value2 = datum.Department;

            r.Offset[4, 0].Value2 = "Specialty:";
            r.Offset[4, 0].Font.Bold = true;
            r.Offset[4, 1].Value2 = datum.Specialty;

            r.Offset[5, 0].Value2 = "User Type:";
            r.Offset[5, 0].Font.Bold = true;
            r.Offset[5, 1].Value2 = datum.UserType;

            r.EntireColumn.NumberFormat = "mm/dd/yyyy";
            r.EntireColumn.ColumnWidth = 16;
            r.Offset[7, 0].Value2 = "End Date";
            r.Offset[7, 0].Font.Bold = true;

            r.Offset[7, 1].EntireColumn.NumberFormat = "0.0";
            r.Offset[7, 1].EntireColumn.ColumnWidth = 12;
            r.Offset[7, 1].Value2 = "Numerator";
            r.Offset[7, 1].Font.Bold = true;

            r.Offset[7, 2].EntireColumn.ColumnWidth = 12;
            r.Offset[7, 2].Value2 = "Denominator";
            r.Offset[7, 2].Font.Bold = true;

            r.Offset[7, 3].EntireColumn.NumberFormat = "0.0";
            r.Offset[7, 3].EntireColumn.ColumnWidth = 20;
            r.Offset[7, 3].Value2 = datum.Metric;
            r.Offset[7, 3].Font.Bold = true;

            Range headings = newSheet.Range[r.Offset[7, 0], r.Offset[7, 3]];
            Borders borders = headings.Borders;
            borders[XlBordersIndex.xlEdgeBottom].Weight = XlBorderWeight.xlThick;

            table.IncrementRow(7);
            tables[datum.ClinicianName] = table;

            return table;
        }

        internal void Import()
        {
            List<string> jsonFiles = AskUserForFiles();
            bool firstFile = true;
            List<NamePlusId> selectedMetrics = new List<NamePlusId>();

            var options = new JsonSerializerOptions();
            options.Converters.Add(new CustomDateTimeConverter());

            foreach (string jsonFile in jsonFiles)
            {
                using (StreamReader r = new StreamReader(jsonFile))
                {
                    application.StatusBar = "Reading " + jsonFile;
                    string json = r.ReadToEnd();

                    Rootobject obj = JsonSerializer.Deserialize<Rootobject>(json, options);

                    // Do we use ALL the metrics or ask the user to select?
                    if (firstFile)
                    {
                        application.StatusBar = "Compiling list of metrics.";
                        selectedMetrics = SelectMetricsToUse(obj.Metrics());

                        // Get sheet names from (perhaps too long) metric names.
                        // Build dictionaries to link metrics to sheets.
                        SheetNamesFromMetrics(selectedMetrics);

                        InitializeLastRowDictionary();

                        firstFile = false;
                    }

                    BuildSheets(obj);
                }
            }
        }

        private void InitializeLastRowDictionary()
        {
            // Sheet name to last row.
            lastRowUsed = new Dictionary<string, int>();

            foreach (NamePlusId metric in metricsToSheets.Keys.ToList())
            {
                lastRowUsed[metric.Combo()] = 0;
            }
        }

        private void InitializeSheet(Worksheet sheet, string metricName)
        {
            Range r = sheet.Cells[1, 1];
            r.Value2 = "Clinician Name";
            r.Font.Bold = true;

            r.Offset[0, 1].Value2 = "Clinician Type";
            r.Offset[0, 1].Font.Bold = true;

            r.Offset[0, 2].Value2 = "Service Area";
            r.Offset[0, 2].Font.Bold = true;

            r.Offset[0, 3].Value2 = "Department";
            r.Offset[0, 3].Font.Bold = true;

            r.Offset[0, 4].Value2 = "Specialty";
            r.Offset[0, 4].Font.Bold = true;

            r.Offset[0, 5].Value2 = "User Type";
            r.Offset[0, 5].Font.Bold = true;

            r.Offset[0, 6].Value2 = "End Date";
            r.Offset[0, 6].Font.Bold = true;
            r.Offset[0, 6].EntireColumn.NumberFormat = "mm/dd/yyyy";
            r.Offset[0, 6].EntireColumn.ColumnWidth = 16;

            r.Offset[0, 7].Value2 = "Numerator";
            r.Offset[0, 7].Font.Bold = true;
            r.Offset[0, 7].EntireColumn.NumberFormat = "0.0";
            r.Offset[0, 7].EntireColumn.ColumnWidth = 12;

            r.Offset[0, 8].Value2 = "Denominator";
            r.Offset[0, 8].Font.Bold = true;
            r.Offset[0, 8].EntireColumn.ColumnWidth = 12;

            r.Offset[0, 9].Value2 = metricName;
            r.Offset[0, 9].Font.Bold = true;
            r.Offset[0, 9].EntireColumn.NumberFormat = "0.0";
            r.Offset[0, 9].EntireColumn.ColumnWidth = 20;

            Range headings = sheet.Range[r, r.Offset[0, 9]];
            Borders borders = headings.Borders;
            borders[XlBordersIndex.xlEdgeBottom].Weight = XlBorderWeight.xlThick;
        }

        private void PutAllPhysiciansOnSameSheet(List<Datum> sortedData)
        {
            foreach (NamePlusId metric in metricsToSheets.Keys.ToList())
            {
                // What was the last row we used on this sheet?
                int lastRow = lastRowUsed[metric.Combo()];

                Worksheet thisSheet;

                // Is this the first time we put data on this sheet?
                if (lastRow == 0)
                {
                    thisSheet = Utilities.CreateNewNamedSheet(metricsToSheets[metric]);
                    InitializeSheet(thisSheet, metric.Combo());
                }
                else
                {
                    // Lookup by name.
                    thisSheet = (Worksheet)thisWorkbook.Worksheets[metricsToSheets[metric]];
                }

                Range r = thisSheet.Cells[1, 1];
                int rowOffset = lastRow + 1;
                int index = 0;

                // Find by metric ID, not just the name.
                List<Datum> dataThisMetric = sortedData.Where(obj => obj.MetricID == metric.id).ToList();

                foreach (Datum datum in dataThisMetric)
                {
                    r.Offset[rowOffset, 0].Value2 = datum.ClinicianName;
                    r.Offset[rowOffset, 1].Value2 = datum.ClinicianType;
                    r.Offset[rowOffset, 2].Value2 = datum.ServiceArea;
                    r.Offset[rowOffset, 3].Value2 = datum.Department;
                    r.Offset[rowOffset, 4].Value2 = datum.Specialty;
                    r.Offset[rowOffset, 5].Value2 = datum.UserType;
                    r.Offset[rowOffset, 6].Value2 = datum.ReportingPeriodEndDate;
                    r.Offset[rowOffset, 7].Value2 = datum.Numerator;
                    r.Offset[rowOffset, 8].Value2 = datum.Denominator;
                    r.Offset[rowOffset, 9].Value2 = datum.Value;
                    index++;

                    if (rowOffset % 100 == 0)
                    {
                        application.StatusBar = "Building sheet for: " + metric.Combo() + " " +
                        index.ToString() + "/" + dataThisMetric.Count;
                    }

                    rowOffset++;
                }

                lastRowUsed[metric.Combo()] = rowOffset - 1;
            }
        }

        private void PutEachPhysicianOnOwnSheet(List<Datum> sortedData)
        {
            foreach (Datum datum in sortedData)
            {
                PhysicianTable table = FindOrCreateTable(datum);
                Range r = table.TopLeft();
                r.Offset[table.Row(), 0].Value2 = datum.ReportingPeriodEndDate;
                r.Offset[table.Row(), 1].Value2 = datum.Numerator;
                r.Offset[table.Row(), 2].Value2 = datum.Denominator;
                r.Offset[table.Row(), 3].Value2 = datum.Value;
                table.IncrementRow();
            }
        }

        // Show the metrics and ask the user to down select.
        private List<NamePlusId> SelectMetricsToUse(List<NamePlusId> metrics)
        {
            List<NamePlusId> selectedMetrics = new List<NamePlusId>();
            application.StatusBar = "Select metrics to extract.";

            // Ask user to downselect.
            using (SelectMetricsForm form = new SelectMetricsForm(metrics))
            {
                if (form.ShowDialog() == DialogResult.OK)
                {
                    selectedMetrics.Clear();
                    selectedMetrics = form.selectedMetrics;
                }
            }

            return selectedMetrics;
        }

        // If the metric names are all long, we can't just truncate to 31 characters
        // and we'll do better by eliminating the common parts.
        // BUT we'll need dictionaries to link them.
        private void SheetNamesFromMetrics(List<NamePlusId> metrics)
        {
            metricsToSheets = new Dictionary<NamePlusId, string>();

            // The default case: sheet names and metric names are the same.
            foreach (NamePlusId metric in metrics)
            {
                metricsToSheets.Add(metric, metric.name);
            }

            // We can just use the names provided if they're short enough to fit on the Sheet tab.
            // (But don't bother if just ONE metric--removing the common part will ERASE the name.)
            if (metrics.Count > 1 && metrics.Any(word => word.name.Length > 31))
            {
                List<string> usableNames = new List<string>();
                List<string> metricNames = metrics.Select(metric => metric.Combo()).ToList();
                string commonPart = Utilities.CommonElements(metricNames);

                if (!string.IsNullOrEmpty(commonPart))
                {
                    metricsToSheets.Clear();

                    foreach (NamePlusId metric in metrics)
                    {
                        string sheetName = metric.Combo().Replace(commonPart, "");
                        metricsToSheets[metric] = sheetName;
                    }
                }
            }
        }
    }
}
