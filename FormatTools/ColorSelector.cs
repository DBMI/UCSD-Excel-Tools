using System;
using System.Drawing;
using System.Windows.Forms;

namespace FormatTools
{
    public static class ColorSelector
    {

        public static Color Select()
        {
            // light blue
            Color selectedColor = Color.FromArgb(221, 235, 247);

            using (ColorDialog colorDialog = new ColorDialog())
            {
                // Allow the user to select custom colors
                colorDialog.AllowFullOpen = true;
                colorDialog.ShowHelp = true;

                // Set the initial color to the dialog.
                colorDialog.Color = RetrieveSavedColor();

                if (colorDialog.ShowDialog() == DialogResult.OK)
                {
                    selectedColor = colorDialog.Color;

                    Properties.Settings.Default.SavedColor = selectedColor.ToArgb().ToString();
                    Properties.Settings.Default.Save();
                }
            }

            return selectedColor;
        }


        public static Color RetrieveSavedColor()
        {
            // light blue
            Color retrievedColor = Color.FromArgb(221, 235, 247);

            string savedColorValue = Properties.Settings.Default.SavedColor;

            if (!string.IsNullOrEmpty(savedColorValue) && int.TryParse(savedColorValue, out int argb))
            {
                retrievedColor = Color.FromArgb(argb);
            }

            return retrievedColor;
        }
    }
}
