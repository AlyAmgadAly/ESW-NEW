using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;

namespace ExcelScore.Classes
{
    public class ColorClass
    {
        private static readonly Color[] colors = new Color[]
        {
            //Color.DarkBlue, // Dark Blue
            Color.Orange, // Orange Accent (Using Orange as a close match)
            //Color.LightBlue, // Aqua
            
            Color.Tan, // Tan
            Color.FromArgb(117, 219, 255),
            Color.FromArgb(241, 182, 163),
            Color.FromArgb(255, 102, 102), // Light Red (custom definition)
            Color.Orange // Orange
            ,Color.YellowGreen
            ,Color.FromArgb(245,215,161)
            ,Color.FromArgb(240,162,142)
            ,Color.FromArgb(186,99,117)

        };
        private static Color lastColor = Color.Empty;
        public Color GenerateRandomColor()
        {
            Random rand = new Random();
            Color nextColor;

            do
            {
                // Select a random color from the array
                nextColor = colors[rand.Next(colors.Length)];
            }
            // Ensure the next color is not the same as the last color generated
            while (nextColor == lastColor);

            // Update the last color with the next color
            lastColor = nextColor;

            return nextColor;
        }

    }
}
