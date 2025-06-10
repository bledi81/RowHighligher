using System;
using System.Windows.Forms;
using System.Drawing;
using System.Text;

namespace RowHighligher
{
    public partial class UnitsConvertHelpForm : Form
    {
        private TabControl tabControl;

        public UnitsConvertHelpForm()
        {
            InitializeComponents();
            this.TopMost = true; // Make help stay on top
        }

        private void InitializeComponents()
        {
            this.Text = "Units Converter Help";
            this.Size = new Size(800, 600);
            this.FormBorderStyle = FormBorderStyle.SizableToolWindow;
            this.StartPosition = FormStartPosition.CenterParent;
            this.MinimumSize = new Size(600, 400);

            // Create custom tab control with light blue selected tab
            tabControl = new TabControl
            {
                Dock = DockStyle.Fill,
                DrawMode = TabDrawMode.OwnerDrawFixed  // Enable custom drawing
            };

            // Add event handlers for custom drawing
            tabControl.DrawItem += new DrawItemEventHandler(TabControl_DrawItem);

            // Overview Tab
            TabPage overviewTab = CreateHelpTab("Overview", @"Units Converter Overview

The Units Converter is a powerful tool for converting between a wide range of measurement units, including those commonly used in the oil, gas, and engineering industries.

Main Features:
• Convert between different units within the same category
• Extensive unit categories: Length, Volume, Pressure, Temperature, Energy, Density, Time, Force, Speed, Flow Rate, Dynamic Viscosity, Kinematic Viscosity, Power, Torque
• Get values directly from Excel cells
• Insert conversion results back to Excel
• Swap units with a single click (swaps only the units, not the values)
• Automatic unit categorization
• Symbol display for easy reference

Key Benefits:
• Simplify complex unit conversions
• Save time on manual calculations
• Reduce errors in unit conversions
• Standardize units across your worksheets
• Seamless Excel integration
• Always accessible from the ribbon

Getting Started:
• Select a category from the dropdown menu
• Enter a value to convert
• Select the original unit (from) and target unit (to)
• Click Convert or press Enter to perform the conversion
• Use Insert Result to place the value in Excel");

            // Basic Usage Tab
            TabPage basicTab = CreateHelpTab("Basic Usage", @"Basic Usage Instructions

Step-by-Step Guide:

1. Select a Category:
   • Choose from: Length, Liquid Volume, Gas Volume, Mass, Pressure, Temperature, Energy, Density, Time, Force, Speed, Flow Rate, Dynamic Viscosity, Kinematic Viscosity, Power, Torque
   • Each category contains relevant units for that measurement type

2. Enter Input Value:
   • Type a number directly in the input field, or
   • Use the 'Get value from cell' button to import from the active Excel cell
   • Press Enter to convert instantly

3. Select Units:
   • From Unit: Select the original unit of your value
   • To Unit: Select the target unit you want to convert to
   • Units display their symbols in brackets for easy identification [m], [ft], etc.

4. Convert:
   • Click the blue 'Convert' button or press Enter to perform the calculation
   • Result appears in the result field

5. Use the Result:
   • View the converted value in the result field
   • Click 'Insert Result' to place the value in the active Excel cell

Tips:
   • Use the '↑↓ Swap Units' button to reverse the conversion direction (swaps only the units)
   • The converter remembers your last category selection
   • Units are organized by industry-standard categories
   • You can detach the converter window to keep it visible while working");

            // Unit Categories Tab
            TabPage categoriesTab = CreateHelpTab("Unit Categories", @"Available Unit Categories

The Units Converter includes the following categories, each with specific units:

1. Length
   • meter [m], kilometer [km], foot [ft], inch [in], mile [mi]

2. Liquid Volume
   • liter [L], cubic meter [m³], US gallon [gal], barrel (oil) [bbl], cubic foot [ft³]

3. Gas Volume
   • standard cubic meter [Sm³], standard cubic foot [scf], thousand standard cubic feet [Mscf], million standard cubic feet [MMscf], billion standard cubic feet [Bscf], normal cubic meter [Nm³], million standard cubic meters [MMSm³]

4. Mass
   • kilogram [kg], metric ton [t], short ton [ton], long ton [long ton], pound [lb]

5. Pressure
   • bar [bar], psi [psi], kPa [kPa], atm [atm]

6. Temperature
   • Celsius [°C], Fahrenheit [°F], Kelvin [K]

7. Energy
   • joule [J], kilojoule [kJ], British Thermal Unit [BTU], therm [thm], kilowatt-hour [kWh]

8. Density
   • kilogram per cubic meter [kg/m³], gram per cubic centimeter [g/cm³], pound per cubic foot [lb/ft³], pound per gallon [lb/gal], API gravity [°API], specific gravity [SG], pound per barrel [lb/bbl]

9. Time
   • second [s], minute [min], hour [h], day [d]

10. Force
   • newton [N], kilonewton [kN], dyne [dyn], kilogram-force [kgf], pound-force [lbf]

11. Speed
   • meter per second [m/s], kilometer per hour [km/h], mile per hour [mph], foot per second [ft/s], knot [kn]

12. Flow Rate
   • cubic meter per second [m³/s], liter per second [L/s], liter per minute [L/min], gallon per minute [gal/min], barrel per day [bbl/d]

13. Dynamic Viscosity
   • pascal second [Pa·s], poise [P], centipoise [cP], pound per foot per second [lb/(ft·s)]

14. Kinematic Viscosity
   • square meter per second [m²/s], stokes [St], centistokes [cSt]

15. Power
   • watt [W], kilowatt [kW], megawatt [MW], horsepower [hp], BTU per hour [BTU/h]

16. Torque
   • newton meter [N·m], kilogram-force meter [kgf·m], pound-force foot [lbf·ft], pound-force inch [lbf·in]");

            // Special Conversions Tab
            TabPage specialTab = CreateHelpTab("Special Conversions", @"Special Conversion Handling

Some unit conversions require special handling due to their non-linear relationships:

Temperature Conversion:
• Temperature conversions don't follow simple multiplication factors
• The converter handles these using specialized formulas:

  Celsius to Fahrenheit: °F = °C × 9/5 + 32
  Fahrenheit to Celsius: °C = (°F - 32) × 5/9
  Celsius to Kelvin: K = °C + 273.15
  Kelvin to Celsius: °C = K - 273.15
  Fahrenheit to Kelvin: K = (°F - 32) × 5/9 + 273.15
  Kelvin to Fahrenheit: °F = (K - 273.15) × 9/5 + 32

Density & API Gravity:
• API gravity is an inverse measure of petroleum liquid density
• Specific formulas are used for these conversions:

  API gravity to kg/m³: ρ = 141.5 / (API + 131.5) × 999.0
  kg/m³ to API gravity: °API = (141.5 / (ρ/999.0)) - 131.5
  
  Specific gravity to kg/m³: ρ = SG × 999.0
  kg/m³ to specific gravity: SG = ρ / 999.0

Notes:
• 999.0 kg/m³ is the reference density of water at 15°C
• Higher API gravity means lower density
• API gravity of water is approximately 10°API");

            // Tips & Tricks Tab
            TabPage tipsTab = CreateHelpTab("Tips & Tricks", @"Tips & Tricks

Excel Integration:
• Place your cursor in an Excel cell before clicking ""Get value from cell"" to import that value
• The ""Insert Result"" button sends the converted value to the current active cell
• You can detach the converter window to keep it visible while you work

Efficiency Tips:
• Use the ""↑↓ Swap Units"" button to quickly reverse your conversion direction
  - This also moves the previous result to the input field
  - Useful for multiple conversions in a sequence

• The converter remembers the last category you selected between uses

• Each category has a default ""from"" and ""to"" unit selection
  - First unit is selected as ""from""
  - Second unit is selected as ""to""

• All unit options show both the name and symbol for easy reference

Error Handling:
• Invalid number inputs will show an error message
• Non-numeric values cannot be inserted into Excel
• Temperature and density conversions use special handling for accurate results

Advanced Usage:
• For frequent conversions, consider setting up Excel formulas using the results
• The converter follows industry-standard conversion factors for all units
• All gas volumes are based on standard temperature and pressure (STP) conditions
• Density conversions properly handle the non-linear API gravity scale");

            // Add tabs in order
            tabControl.TabPages.AddRange(new TabPage[] { 
                overviewTab,
                basicTab, 
                categoriesTab,
                specialTab,
                tipsTab
            });

            this.Controls.Add(tabControl);
        }

        private void TabControl_DrawItem(object sender, DrawItemEventArgs e)
        {
            TabPage page = tabControl.TabPages[e.Index];
            Rectangle bounds = tabControl.GetTabRect(e.Index);
            
            // Define colors
            Color selectedColor = Color.LightSkyBlue;
            Color unselectedColor = SystemColors.Control;
            Color textColor = Color.Black;

            // Fill the background
            using (SolidBrush brush = new SolidBrush(e.Index == tabControl.SelectedIndex ? selectedColor : unselectedColor))
            {
                e.Graphics.FillRectangle(brush, bounds);
            }

            // Draw the text
            StringFormat stringFlags = new StringFormat
            {
                Alignment = StringAlignment.Center,
                LineAlignment = StringAlignment.Center
            };

            using (SolidBrush brush = new SolidBrush(textColor))
            {
                e.Graphics.DrawString(page.Text, this.Font, brush, bounds, stringFlags);
            }
        }

        private RichTextBox CreateHelpTextBox(string content)
        {
            RichTextBox textBox = new RichTextBox
            {
                Dock = DockStyle.Fill,
                ReadOnly = true,
                BackColor = SystemColors.Window,
                Font = new Font("Segoe UI", 10),
                Multiline = true,
                AcceptsTab = true,
                WordWrap = true,
                ScrollBars = RichTextBoxScrollBars.Vertical,
                DetectUrls = false // Prevent automatic URL detection which can affect formatting
            };

            // Use the RTF parser to properly handle special characters
            textBox.Text = content.Replace("•", "•"); // Ensure bullet points are consistent

            return textBox;
        }

        private TabPage CreateHelpTab(string title, string content)
        {
            var tab = new TabPage(title);
            tab.Controls.Add(CreateHelpTextBox(content));
            return tab;
        }
    }
}