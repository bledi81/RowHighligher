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

The Units Converter is a powerful tool for converting between different measurement units commonly used in the oil and gas industry.

Main Features:
• Convert between different units within the same category
• Multiple unit categories: Length, Volume, Pressure, Temperature, etc.
• Get values directly from Excel cells
• Insert conversion results back to Excel
• Swap units with a single click
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
• Click Convert to perform the conversion
• Use Insert Result to place the value in Excel");

            // Basic Usage Tab
            TabPage basicTab = CreateHelpTab("Basic Usage", @"Basic Usage Instructions

Step-by-Step Guide:

1. Select a Category:
   • Choose from Length, Liquid Volume, Gas Volume, Mass, Pressure, Temperature, Energy, or Density
   • Each category contains relevant units for that measurement type

2. Enter Input Value:
   • Type a number directly in the input field, or
   • Use the ""Get value from cell"" button to import from the active Excel cell

3. Select Units:
   • From Unit: Select the original unit of your value
   • To Unit: Select the target unit you want to convert to
   • Units display their symbols in brackets for easy identification [m], [ft], etc.

4. Convert:
   • Click the blue ""Convert"" button to perform the calculation
   • Result appears in the result field

5. Use the Result:
   • View the converted value in the result field
   • Click ""Insert Result"" to place the value in the active Excel cell

Tips:
   • Use the ""↑↓ Swap Units"" button to reverse the conversion direction
   • The converter remembers your last category selection
   • Units are organized by industry-standard categories
   • You can detach the converter window to keep it visible while working");

            // Unit Categories Tab
            TabPage categoriesTab = CreateHelpTab("Unit Categories", @"Available Unit Categories

The Units Converter includes the following categories, each with specific units relevant to oil and gas operations:

1. Length
   • meter [m] - Base SI unit for length
   • kilometer [km] - 1000 meters
   • foot [ft] - 0.3048 meters
   • inch [in] - 0.0254 meters
   • mile [mi] - 1609.344 meters

2. Liquid Volume
   • liter [L] - Base unit for liquid volume conversion
   • cubic meter [m³] - 1000 liters
   • US gallon [gal] - 3.78541 liters
   • barrel (oil) [bbl] - 158.987 liters
   • cubic foot [ft³] - 28.3168 liters

3. Gas Volume
   • standard cubic meter [Sm³] - Base unit for gas volume
   • standard cubic foot [scf] - 0.0283168 Sm³
   • thousand standard cubic feet [Mscf] - 28.3168 Sm³
   • million standard cubic feet [MMscf] - 28316.8 Sm³
   • billion standard cubic feet [Bscf] - 28316800.0 Sm³
   • normal cubic meter [Nm³] - At normal conditions
   • million standard cubic meters [MMSm³] - 1000000.0 Sm³

4. Mass
   • kilogram [kg] - Base SI unit for mass
   • metric ton [t] - 1000 kg
   • short ton [ton] - 907.185 kg
   • long ton [long ton] - 1016.05 kg
   • pound [lb] - 0.453592 kg

5. Pressure
   • bar [bar] - Base unit for pressure conversion
   • psi [psi] - 0.0689476 bar
   • kPa [kPa] - 0.01 bar
   • atm [atm] - 1.01325 bar

6. Temperature
   • Celsius [°C] - Base unit for temperature conversion
   • Fahrenheit [°F] - (°F = °C × 9/5 + 32)
   • Kelvin [K] - (K = °C + 273.15)

7. Energy
   • joule [J] - Base SI unit for energy
   • kilojoule [kJ] - 1000 joules
   • British Thermal Unit [BTU] - 1055.06 joules
   • therm [thm] - 105506000.0 joules
   • kilowatt-hour [kWh] - 3600000.0 joules

8. Density
   • kilogram per cubic meter [kg/m³] - Base SI unit for density
   • gram per cubic centimeter [g/cm³] - 1000.0 kg/m³
   • pound per cubic foot [lb/ft³] - 16.0185 kg/m³
   • pound per gallon [lb/gal] - 119.8264 kg/m³
   • API gravity [°API] - Special oil industry measure
   • specific gravity [SG] - Ratio relative to water density
   • pound per barrel [lb/bbl] - Oil industry measure");

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
            // Create a RichTextBox with proper Unicode support
            RichTextBox textBox = new RichTextBox
            {
                Dock = DockStyle.Fill,
                ReadOnly = true,
                BackColor = SystemColors.Window,
                Font = new Font("Segoe UI", 10),
                Multiline = true,
                AcceptsTab = true,
                WordWrap = true,
                ScrollBars = RichTextBoxScrollBars.Vertical
            };

            // Set the text with explicit encoding handling
            textBox.Text = content;
            
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