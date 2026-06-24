using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Windows.Forms;

namespace ListTools
{
    public partial class SelectMetricsForm : Form
    {
        public List<NamePlusId> selectedMetrics;
        public Dictionary<string, NamePlusId> metricsFromNames;

        public SelectMetricsForm(List<NamePlusId> metrics)
        {
            InitializeComponent();
            PopulateDictionary(metrics);
            Utilities.PopulateListBox(metricsListBox, metricsFromNames.Keys.ToList());
            selectedMetrics = new List<NamePlusId>();  
        }

        private void cancelButton_Click(object sender, EventArgs e)
        {
            DialogResult = DialogResult.Cancel;
            Close();
        }

        private void okButton_Click(object sender, EventArgs e)
        {
            selectedMetrics.Clear();

            List<string> selectedNames = metricsListBox.SelectedItems.Cast<string>().ToList();

            foreach (string name in selectedNames)
            {
                selectedMetrics.Add(metricsFromNames[name]);
            }

            DialogResult = DialogResult.OK;
            Close();
        }

        private void PopulateDictionary(List<NamePlusId> metrics)
        {
            metricsFromNames = new Dictionary<string, NamePlusId>();

            foreach(NamePlusId metric in metrics)
            {
                if (!metricsFromNames.ContainsKey(metric.Combo()))
                {
                    metricsFromNames[metric.Combo()] = metric;
                }                
            }
        }
    }
}
