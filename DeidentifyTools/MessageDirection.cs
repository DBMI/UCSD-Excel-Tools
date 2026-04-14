using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace DeidentifyTools
{
    public enum MessageDirectionEnum
    {
        FromPatient,
        ToPatient,
        None
    }

    internal class MessageDirection
    {
        private MessageDirectionEnum _direction;

        internal MessageDirection(string columnName)
        {
            _direction = MessageDirectionEnum.None;

            if (!string.IsNullOrEmpty(columnName))
            {
                if (columnName.ToLower().Contains("from patient"))
                {
                    _direction = MessageDirectionEnum.FromPatient;
                }
                else if (columnName.ToLower().Contains("to patient"))
                {
                    _direction = MessageDirectionEnum.ToPatient;
                }
                else
                {
                    // If we can't figure it out from the column name, ask user directly;
                    using (MessageDirectionForm form = new MessageDirectionForm())
                    {
                        var result = form.ShowDialog();

                        if (result == DialogResult.OK)
                        {
                            _direction = form.direction;
                        }
                    }
                }
            }
        }

        internal MessageDirectionEnum Direction()
        {
            return _direction;
        }
    }
}
