using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace DeidentifyTools
{
    public partial class MessageDirectionForm : Form
    {
        public MessageDirectionEnum direction = MessageDirectionEnum.None;

        public MessageDirectionForm()
        {
            InitializeComponent();
        }

        private void fromPatientRadioButton_Clicked(object sender, EventArgs e)
        {
            if (fromPatientRadioButton.Checked)
            {
                direction = MessageDirectionEnum.FromPatient;
                toPatientRadioButton.Checked = false;
                noneRadioButton.Checked = false;
            }
        }

        private void toPatientRadioButton_Click(object sender, EventArgs e)
        {
            if (toPatientRadioButton.Checked)
            {
                direction = MessageDirectionEnum.ToPatient;
                fromPatientRadioButton.Checked = false;
                noneRadioButton.Checked = false;
            }
        }

        private void noneRadioButton_Click(object sender, EventArgs e)
        {
            if (noneRadioButton.Checked) 
            {
                direction = MessageDirectionEnum.None;
                fromPatientRadioButton.Checked = false;
                toPatientRadioButton.Checked= false;
            }
        }

        private void okButton_Click(object sender, EventArgs e)
        {
            DialogResult = DialogResult.OK;
            Close();
        }

        private void cancelButton_Click(object sender, EventArgs e)
        {
            DialogResult = DialogResult.Cancel;
            Close();
        }
    }
}
