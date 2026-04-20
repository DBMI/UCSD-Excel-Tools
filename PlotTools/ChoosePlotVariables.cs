using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;

namespace PlotTools
{
    public partial class ChoosePlotVariables : Form
    {
        public string xColumnName = string.Empty;
        public string yColumnName = string.Empty;

        private bool disableCallbacks;
        private bool initializing;

        public ChoosePlotVariables(List<string> columnNames)
        {
            InitializeComponent();

            disableCallbacks = false;
            initializing = true;
            Utilities.PopulateListBox(xListBox, columnNames, enableWhenPopulated: true);
            Utilities.PopulateListBox(yListBox, columnNames, enableWhenPopulated: true);
            initializing = false;
        }

        public void CancelButton_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.Cancel;
            this.Close();
        }

        private void EnableRunWhenReady()
        {
            if (initializing)
            {
                return;
            }

            okButton.Enabled =
                    !string.IsNullOrEmpty(xColumnName) && !string.IsNullOrEmpty(yColumnName);
        }

        public void OkButton_Click(object sender, EventArgs e)
        {
            xColumnName = xListBox.SelectedItem.ToString();
            yColumnName = yListBox.SelectedItem.ToString();

            this.DialogResult = DialogResult.OK;
            this.Close();
        }

        private void xListBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (initializing || disableCallbacks)
            {
                return;
            }

            disableCallbacks = true;

            xColumnName = xListBox.SelectedItem.ToString();

            disableCallbacks = false;

            EnableRunWhenReady();
        }

        private void yListBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (initializing || disableCallbacks)
            {
                return;
            }

            disableCallbacks = true;

            yColumnName = yListBox.SelectedItem.ToString();

            disableCallbacks = false;

            EnableRunWhenReady();
        }
    }
}
