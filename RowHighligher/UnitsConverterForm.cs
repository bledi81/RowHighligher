using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;

namespace RowHighligher
{
    public partial class UnitsConverterForm : Form
    {
        private TextBox inputTextBox;
        private ComboBox fromUnitComboBox;
        private ComboBox toUnitComboBox;
        private Button convertButton;
        private TextBox resultTextBox;
        private Button insertButton;
        private Button getFromCellButton;
        private Button helpButton;
        private Label inputLabel;
        private Label resultLabel;
        private ComboBox categoryComboBox;
        private PictureBox logoPictureBox;
        private string[] categoryNames;

        private static readonly List<UnitCategory> Categories = UnitConverter.GetOilGasCategories();

        private Panel customTitleBar;
        private Label titleLabel;
        private Button closeButton;
        private Button minimizeButton;
        private Point dragOffset;
        private bool dragging = false;

        public UnitsConverterForm()
        {
            InitializeComponent();
            if (categoryComboBox.Items.Count > 0)
                categoryComboBox.SelectedIndex = 0;
            TryLoadFromExcel();
            this.TopMost = Properties.Settings.Default.IsConverterDetached;
        }

        private void InitializeComponent()
        {
            this.Text = "Units Converter";
            this.Size = new Size(500, 330);
            
            // Allow resizing
            this.FormBorderStyle = FormBorderStyle.None;
            
            // Set the current size as the minimum size
            this.MinimumSize = new Size(600, 330);
            
            // Allow minimizing and maximizing
            // this.MaximizeBox = true;
            // this.MinimizeBox = true;
            
            this.StartPosition = FormStartPosition.CenterParent;
            this.TopMost = true;

            // Prepare categories
            categoryNames = new string[Categories.Count];
            for (int i = 0; i < Categories.Count; i++) categoryNames[i] = Categories[i].Name;

            // Custom title bar
            customTitleBar = new Panel
            {
                Dock = DockStyle.Fill,
                Height = 36,
                BackColor = Color.Orange
            };
            titleLabel = new Label
            {
                Text = "Units Converter",
                Dock = DockStyle.Left,
                AutoSize = false,
                TextAlign = ContentAlignment.MiddleLeft,
                Font = new Font("Segoe UI", 11, FontStyle.Bold),
                ForeColor = Color.Black,
                Padding = new Padding(10, 0, 0, 0),
                Width = 220
            };
            closeButton = new Button
            {
                Text = "✕",
                Dock = DockStyle.Right,
                Width = 36,
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.Orange,
                ForeColor = Color.Black,
                Font = new Font("Segoe UI", 11, FontStyle.Bold),
                TabStop = false
            };
            closeButton.FlatAppearance.BorderSize = 0;
            closeButton.Click += (s, e) => this.Close();
            minimizeButton = new Button
            {
                Text = "_",
                Dock = DockStyle.Right,
                Width = 36,
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.Orange,
                ForeColor = Color.Black,
                Font = new Font("Segoe UI", 11, FontStyle.Bold),
                TabStop = false
            };
            minimizeButton.FlatAppearance.BorderSize = 0;
            minimizeButton.Click += (s, e) => this.WindowState = FormWindowState.Minimized;
            customTitleBar.Controls.Add(closeButton);
            customTitleBar.Controls.Add(minimizeButton);
            customTitleBar.Controls.Add(titleLabel);
            customTitleBar.MouseDown += CustomTitleBar_MouseDown;
            customTitleBar.MouseMove += CustomTitleBar_MouseMove;
            customTitleBar.MouseUp += CustomTitleBar_MouseUp;
            titleLabel.MouseDown += CustomTitleBar_MouseDown;
            titleLabel.MouseMove += CustomTitleBar_MouseMove;
            titleLabel.MouseUp += CustomTitleBar_MouseUp;

            // Main layout: 1 column (controls)
            var mainPanel = new TableLayoutPanel { Dock = DockStyle.Fill, RowCount = 1, ColumnCount = 1, Padding = new Padding(10) };

            // Controls panel (right)
            var controlsPanel = new TableLayoutPanel { Dock = DockStyle.Fill, RowCount = 6, ColumnCount = 3 };
            
            // Use percentage-based column styles
            controlsPanel.ColumnStyles.Clear();
            controlsPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 100)); // Label/input/result (keep fixed width for labels)
            controlsPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 40)); // Main input/result column
            controlsPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 60)); // Third column
            
            // Use percentage-based row styles instead of absolute
            controlsPanel.RowStyles.Add(new RowStyle(SizeType.Percent, 16)); // Category (~40px at default size)
            controlsPanel.RowStyles.Add(new RowStyle(SizeType.Percent, 16)); // Input
            controlsPanel.RowStyles.Add(new RowStyle(SizeType.Percent, 16)); // Get from cell
            controlsPanel.RowStyles.Add(new RowStyle(SizeType.Percent, 22)); // Convert (slightly larger)
            controlsPanel.RowStyles.Add(new RowStyle(SizeType.Percent, 16)); // Result
            controlsPanel.RowStyles.Add(new RowStyle(SizeType.Percent, 16)); // Insert

            var smallFont = new Font("Segoe UI", 11, FontStyle.Bold);
            var smallFontNormal = new Font("Segoe UI", 11, FontStyle.Regular);

            // Category row (top)
            var categoryLabel = new Label { Text = "Category:", Anchor = AnchorStyles.Right, AutoSize = true, Font = smallFontNormal };
            categoryComboBox = new ComboBox { Dock = DockStyle.Fill, DropDownStyle = ComboBoxStyle.DropDownList, Font = smallFontNormal };
            categoryComboBox.Items.AddRange(categoryNames);
            categoryComboBox.SelectedIndexChanged += CategoryComboBox_SelectedIndexChanged;
            logoPictureBox = new PictureBox
            {
                Image = Properties.Resources.Bankers_Logo_Albania,
                SizeMode = PictureBoxSizeMode.Zoom,
                Dock = DockStyle.Fill,
                Margin = new Padding(0, 0, 0, 0)
            };
            controlsPanel.Controls.Add(categoryLabel, 0, 0);
            controlsPanel.Controls.Add(categoryComboBox, 1, 0);
            controlsPanel.Controls.Add(logoPictureBox, 2, 0);

            // Input row
            inputLabel = new Label { Text = "Input Value:", Anchor = AnchorStyles.Right, AutoSize = true, Font = smallFontNormal };
            inputTextBox = new TextBox { Dock = DockStyle.Fill, Font = smallFontNormal };
            inputTextBox.KeyDown += InputTextBox_KeyDown;
            fromUnitComboBox = new ComboBox { Dock = DockStyle.Fill, DropDownStyle = ComboBoxStyle.DropDownList, Font = smallFontNormal };
            controlsPanel.Controls.Add(inputLabel, 0, 1);
            controlsPanel.Controls.Add(inputTextBox, 1, 1);
            controlsPanel.Controls.Add(fromUnitComboBox, 2, 1);

            // Get from cell row
            getFromCellButton = new Button { Text = "Get value from cell", Dock = DockStyle.Fill, Font = smallFont };
            getFromCellButton.Click += GetFromCellButton_Click;
            controlsPanel.Controls.Add(getFromCellButton, 2, 2);

            // Convert row
            convertButton = new Button { Text = "Convert", Dock = DockStyle.Fill, BackColor = Color.LightSkyBlue, Font = smallFont };
            convertButton.Click += ConvertButton_Click;
            controlsPanel.Controls.Add(convertButton, 1, 3);

            // Add the Swap Units button in the same row with proper Unicode arrows
            Button swapUnitsButton = new Button 
            { 
                Text = "↑↓ Swap Units", 
                Dock = DockStyle.Fill, 
                Font = smallFont
            };
            swapUnitsButton.Click += SwapUnitsButton_Click;
            controlsPanel.Controls.Add(swapUnitsButton, 2, 3);

            // Result row
            resultLabel = new Label { Text = "Result:", Anchor = AnchorStyles.Right, AutoSize = true, Font = smallFontNormal };
            resultTextBox = new TextBox { Dock = DockStyle.Fill, ReadOnly = true, Font = smallFontNormal };
            toUnitComboBox = new ComboBox { Dock = DockStyle.Fill, DropDownStyle = ComboBoxStyle.DropDownList, Font = smallFontNormal };
            controlsPanel.Controls.Add(resultLabel, 0, 4);
            controlsPanel.Controls.Add(resultTextBox, 1, 4);
            controlsPanel.Controls.Add(toUnitComboBox, 2, 4);

            // Insert row
            insertButton = new Button { Text = "Insert Result", Dock = DockStyle.Fill, Font = smallFont };
            insertButton.Click += InsertButton_Click;
            controlsPanel.Controls.Add(insertButton, 2, 5);

            // Help button (new)
            helpButton = new Button { Text = "Help", Dock = DockStyle.Fill, Font = smallFont };
            helpButton.Click += HelpButton_Click;
            controlsPanel.Controls.Add(helpButton, 0, 5);

            mainPanel.Controls.Add(controlsPanel, 0, 0);

            // Root layout: 2 rows, title bar and mainPanel
            var rootPanel = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                RowCount = 2,
                ColumnCount = 1
            };
            rootPanel.RowStyles.Add(new RowStyle(SizeType.Absolute, 36)); // Title bar height
            rootPanel.RowStyles.Add(new RowStyle(SizeType.Percent, 100)); // Main UI
            rootPanel.Controls.Add(customTitleBar, 0, 0);
            rootPanel.Controls.Add(mainPanel, 0, 1);
            this.Controls.Clear();
            this.Controls.Add(rootPanel);
        }

        protected override void OnActivated(EventArgs e)
        {
            base.OnActivated(e);
            customTitleBar.BackColor = Color.Orange;
            closeButton.BackColor = Color.Orange;
            minimizeButton.BackColor = Color.Orange;
        }

        protected override void OnDeactivate(EventArgs e)
        {
            base.OnDeactivate(e);
            customTitleBar.BackColor = Color.FromArgb(255, 230, 180); // Lighter orange when inactive
            closeButton.BackColor = Color.FromArgb(255, 230, 180);
            minimizeButton.BackColor = Color.FromArgb(255, 230, 180);
        }

        private void CustomTitleBar_MouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                dragging = true;
                dragOffset = new Point(e.X, e.Y);
            }
        }

        private void CustomTitleBar_MouseMove(object sender, MouseEventArgs e)
        {
            if (dragging)
            {
                Point p = PointToScreen(e.Location);
                this.Location = new Point(p.X - dragOffset.X, p.Y - dragOffset.Y);
            }
        }

        private void CustomTitleBar_MouseUp(object sender, MouseEventArgs e)
        {
            dragging = false;
        }

        private void HelpButton_Click(object sender, EventArgs e)
        {
            // Create and show the help form
            using (var helpForm = new UnitsConvertHelpForm())
            {
                helpForm.ShowDialog(this);
            }
        }

        private void LoadUnits()
        {
            // Only populate if a category is selected
            if (categoryComboBox == null || categoryComboBox.SelectedIndex < 0) return;
            var cat = Categories[categoryComboBox.SelectedIndex];
            fromUnitComboBox.Items.Clear();
            toUnitComboBox.Items.Clear();
            
            foreach (var unitName in cat.Units.Keys)
            {
                string symbol = cat.UnitSymbols[unitName];
                var unitItem = new UnitItem(unitName, symbol);
                fromUnitComboBox.Items.Add(unitItem);
                toUnitComboBox.Items.Add(unitItem);
            }
            
            if (fromUnitComboBox.Items.Count > 0) fromUnitComboBox.SelectedIndex = 0;
            if (toUnitComboBox.Items.Count > 1) toUnitComboBox.SelectedIndex = 1;
        }

        private void TryLoadFromExcel()
        {
            try
            {
                Excel.Application excelApp = Globals.ThisAddIn.Application;
                if (excelApp != null)
                {
                    Excel.Range activeCell = excelApp.ActiveCell;
                    if (activeCell != null && activeCell.Value2 != null)
                    {
                        inputTextBox.Text = activeCell.Value2.ToString();
                        Marshal.ReleaseComObject(activeCell);
                    }
                }
            }
            catch { }
        }

        private void ConvertButton_Click(object sender, EventArgs e)
        {
            PerformConversion();
        }

        private void InsertButton_Click(object sender, EventArgs e)
        {
            try
            {
                Excel.Application excelApp = Globals.ThisAddIn.Application;
                if (excelApp != null)
                {
                    Excel.Range activeCell = excelApp.ActiveCell;
                    if (activeCell != null)
                    {
                        double value;
                        if (double.TryParse(resultTextBox.Text, out value))
                        {
                            activeCell.Value = value;
                            Marshal.ReleaseComObject(activeCell);
                        }
                        else
                        {
                            MessageBox.Show("Cannot insert non-numeric value", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error inserting value: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void GetFromCellButton_Click(object sender, EventArgs e)
        {
            TryLoadFromExcel();
        }

        private void CategoryComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            LoadUnits();
        }

        private void SwapUnitsButton_Click(object sender, EventArgs e)
        {
            // Store current selections
            object fromUnit = fromUnitComboBox.SelectedItem;
            object toUnit = toUnitComboBox.SelectedItem;
            
            // Skip if any unit is not selected
            if (fromUnit == null || toUnit == null)
                return;
            
            // Swap the units only
            int fromIndex = fromUnitComboBox.SelectedIndex;
            int toIndex = toUnitComboBox.SelectedIndex;
            fromUnitComboBox.SelectedIndex = toIndex;
            toUnitComboBox.SelectedIndex = fromIndex;
        }

        // Extract the conversion logic into its own method so it can be reused
        private void PerformConversion()
        {
            double inputValue;
            if (!double.TryParse(inputTextBox.Text, out inputValue))
            {
                MessageBox.Show("Invalid input value.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            
            if (!(fromUnitComboBox.SelectedItem is UnitItem fromUnitItem) || 
                !(toUnitComboBox.SelectedItem is UnitItem toUnitItem))
            {
                MessageBox.Show("Please select both units.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            
            string fromUnit = fromUnitItem.Name;
            string toUnit = toUnitItem.Name;
            
            try
            {
                double result = UnitConverter.Convert(inputValue, fromUnit, toUnit);
                resultTextBox.Text = result.ToString("G10");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Conversion Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void InputTextBox_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                ConvertButton_Click(convertButton, EventArgs.Empty);
                e.Handled = true;
                e.SuppressKeyPress = true;
            }
        }
    }

    public static class UnitConverter
    {
        // Oil & Gas relevant units
        public static List<UnitCategory> GetOilGasCategories()
        {
            return new List<UnitCategory>
            {
                new UnitCategory("Length", 
                    new Dictionary<string, double>
                    {
                        {"meter", 1.0},
                        {"kilometer", 1000.0},
                        {"foot", 0.3048},
                        {"inch", 0.0254},
                        {"mile", 1609.344}
                    },
                    new Dictionary<string, string>
                    {
                        {"meter", "m"},
                        {"kilometer", "km"},
                        {"foot", "ft"},
                        {"inch", "in"},
                        {"mile", "mi"}
                    }),
                new UnitCategory("Liquid Volume", 
                    new Dictionary<string, double>
                    {
                        {"liter", 1.0},
                        {"cubic meter", 1000.0},
                        {"US gallon", 3.78541},
                        {"barrel (oil)", 158.987},
                        {"cubic foot", 28.3168}
                    },
                    new Dictionary<string, string>
                    {
                        {"liter", "L"},
                        {"cubic meter", "m³"},
                        {"US gallon", "gal"},
                        {"barrel (oil)", "bbl"},
                        {"cubic foot", "ft³"}
                    }),
                new UnitCategory("Gas Volume", 
                    new Dictionary<string, double>
                    {
                        {"standard cubic meter", 1.0},
                        {"standard cubic foot", 0.0283168},
                        {"thousand standard cubic feet", 28.3168},
                        {"million standard cubic feet", 28316.8},
                        {"billion standard cubic feet", 28316800.0},
                        {"normal cubic meter", 1.0},
                        {"million standard cubic meters", 1000000.0}
                    },
                    new Dictionary<string, string>
                    {
                        {"standard cubic meter", "Sm³"},
                        {"standard cubic foot", "scf"},
                        {"thousand standard cubic feet", "Mscf"},
                        {"million standard cubic feet", "MMscf"},
                        {"billion standard cubic feet", "Bscf"},
                        {"normal cubic meter", "Nm³"},
                        {"million standard cubic meters", "MMSm³"}
                    }),
                new UnitCategory("Mass", 
                    new Dictionary<string, double>
                    {
                        {"kilogram", 1.0},
                        {"metric ton", 1000.0},
                        {"short ton", 907.185},
                        {"long ton", 1016.05},
                        {"pound", 0.453592}
                    },
                    new Dictionary<string, string>
                    {
                        {"kilogram", "kg"},
                        {"metric ton", "t"},
                        {"short ton", "ton"},
                        {"long ton", "long ton"},
                        {"pound", "lb"}
                    }),
                new UnitCategory("Pressure", 
                    new Dictionary<string, double>
                    {
                        {"bar", 1.0},
                        {"psi", 0.0689476},
                        {"kPa", 0.01},
                        {"atm", 1.01325}
                    },
                    new Dictionary<string, string>
                    {
                        {"bar", "bar"},
                        {"psi", "psi"},
                        {"kPa", "kPa"},
                        {"atm", "atm"}
                    }),
                new UnitCategory("Temperature", 
                    new Dictionary<string, double>
                    {
                        {"Celsius", 1.0},
                        {"Fahrenheit", double.NaN}, // handled specially
                        {"Kelvin", double.NaN} // handled specially
                    },
                    new Dictionary<string, string>
                    {
                        {"Celsius", "°C"},
                        {"Fahrenheit", "°F"},
                        {"Kelvin", "K"}
                    }),
                new UnitCategory("Energy", 
                    new Dictionary<string, double>
                    {
                        {"joule", 1.0},
                        {"kilojoule", 1000.0},
                        {"British Thermal Unit", 1055.06},
                        {"therm", 105506000.0},
                        {"kilowatt-hour", 3600000.0}
                    },
                    new Dictionary<string, string>
                    {
                        {"joule", "J"},
                        {"kilojoule", "kJ"},
                        {"British Thermal Unit", "BTU"},
                        {"therm", "thm"},
                        {"kilowatt-hour", "kWh"}
                    }),
                new UnitCategory("Density", 
                    new Dictionary<string, double>
                    {
                        {"kilogram per cubic meter", 1.0},        // Base unit (SI)
                        {"gram per cubic centimeter", 1000.0},    // Same as kg/L or g/mL
                        {"pound per cubic foot", 16.0185},        // lb/ft³
                        {"pound per gallon", 119.8264},           // lb/gal (US)
                        {"API gravity", double.NaN},              // Special handling required
                        {"specific gravity", double.NaN},         // Special handling required
                        {"pound per barrel", 0.158987 * 119.8264} // lb/bbl
                    },
                    new Dictionary<string, string>
                    {
                        {"kilogram per cubic meter", "kg/m³"},
                        {"gram per cubic centimeter", "g/cm³"},
                        {"pound per cubic foot", "lb/ft³"},
                        {"pound per gallon", "lb/gal"},
                        {"API gravity", "°API"},
                        {"specific gravity", "SG"},
                        {"pound per barrel", "lb/bbl"}
                    }),
                new UnitCategory("Time",
                    new Dictionary<string, double>
                    {
                        {"second", 1.0},
                        {"minute", 60.0},
                        {"hour", 3600.0},
                        {"day", 86400.0}
                    },
                    new Dictionary<string, string>
                    {
                        {"second", "s"},
                        {"minute", "min"},
                        {"hour", "h"},
                        {"day", "d"}
                    }),
                new UnitCategory("Force",
                    new Dictionary<string, double>
                    {
                        {"newton", 1.0},
                        {"kilonewton", 1000.0},
                        {"dyne", 1e-5},
                        {"kilogram-force", 9.80665},
                        {"pound-force", 4.4482216}
                    },
                    new Dictionary<string, string>
                    {
                        {"newton", "N"},
                        {"kilonewton", "kN"},
                        {"dyne", "dyn"},
                        {"kilogram-force", "kgf"},
                        {"pound-force", "lbf"}
                    }),
                new UnitCategory("Speed",
                    new Dictionary<string, double>
                    {
                        {"meter per second", 1.0},
                        {"kilometer per hour", 0.277777778},
                        {"mile per hour", 0.44704},
                        {"foot per second", 0.3048},
                        {"knot", 0.514444}
                    },
                    new Dictionary<string, string>
                    {
                        {"meter per second", "m/s"},
                        {"kilometer per hour", "km/h"},
                        {"mile per hour", "mph"},
                        {"foot per second", "ft/s"},
                        {"knot", "kn"}
                    }),
                new UnitCategory("Flow Rate",
                    new Dictionary<string, double>
                    {
                        {"cubic meter per second", 1.0},
                        {"liter per second", 0.001},
                        {"liter per minute", 0.001/60.0},
                        {"gallon per minute", 0.00378541/60.0},
                        {"barrel per day", 0.158987/86400.0}
                    },
                    new Dictionary<string, string>
                    {
                        {"cubic meter per second", "m³/s"},
                        {"liter per second", "L/s"},
                        {"liter per minute", "L/min"},
                        {"gallon per minute", "gal/min"},
                        {"barrel per day", "bbl/d"}
                    }),
                new UnitCategory("Dynamic Viscosity",
                    new Dictionary<string, double>
                    {
                        {"pascal second", 1.0},
                        {"poise", 0.1},
                        {"centipoise", 0.001},
                        {"pound per foot per second", 1.48816}
                    },
                    new Dictionary<string, string>
                    {
                        {"pascal second", "Pa·s"},
                        {"poise", "P"},
                        {"centipoise", "cP"},
                        {"pound per foot per second", "lb/(ft·s)"}
                    }),
                new UnitCategory("Kinematic Viscosity",
                    new Dictionary<string, double>
                    {
                        {"square meter per second", 1.0},
                        {"stokes", 0.0001},
                        {"centistokes", 0.000001}
                    },
                    new Dictionary<string, string>
                    {
                        {"square meter per second", "m²/s"},
                        {"stokes", "St"},
                        {"centistokes", "cSt"}
                    }),
                new UnitCategory("Power",
                    new Dictionary<string, double>
                    {
                        {"watt", 1.0},
                        {"kilowatt", 1000.0},
                        {"megawatt", 1000000.0},
                        {"horsepower", 745.699872},
                        {"BTU per hour", 0.29307107}
                    },
                    new Dictionary<string, string>
                    {
                        {"watt", "W"},
                        {"kilowatt", "kW"},
                        {"megawatt", "MW"},
                        {"horsepower", "hp"},
                        {"BTU per hour", "BTU/h"}
                    }),
                new UnitCategory("Torque",
                    new Dictionary<string, double>
                    {
                        {"newton meter", 1.0},
                        {"kilogram-force meter", 9.80665},
                        {"pound-force foot", 1.355817948},
                        {"pound-force inch", 0.112984829}
                    },
                    new Dictionary<string, string>
                    {
                        {"newton meter", "N·m"},
                        {"kilogram-force meter", "kgf·m"},
                        {"pound-force foot", "lbf·ft"},
                        {"pound-force inch", "lbf·in"}
                    })
            };
        }

        public static double Convert(double value, string fromUnit, string toUnit)
        {
            // Find category
            foreach (var cat in GetOilGasCategories())
            {
                if (cat.Units.ContainsKey(fromUnit) && cat.Units.ContainsKey(toUnit))
                {
                    if (cat.Name == "Temperature")
                        return ConvertTemperature(value, fromUnit, toUnit);
                    else if (cat.Name == "Density" && (fromUnit == "API gravity" || toUnit == "API gravity" || 
                            fromUnit == "specific gravity" || toUnit == "specific gravity"))
                        return ConvertDensity(value, fromUnit, toUnit);
                    
                    // Convert to base (SI) then to target
                    double baseValue = value * cat.Units[fromUnit];
                    return baseValue / cat.Units[toUnit];
                }
            }
            throw new Exception($"Cannot convert from {fromUnit} to {toUnit}.");
        }

        private static double ConvertTemperature(double value, string from, string to)
        {
            if (from == to) return value;
            // Convert from -> Celsius
            double celsius = from == "Celsius" ? value :
                             from == "Fahrenheit" ? (value - 32) * 5.0 / 9.0 :
                             from == "Kelvin" ? value - 273.15 :
                             throw new Exception("Unknown temperature unit");
            // Celsius -> to
            if (to == "Celsius") return celsius;
            if (to == "Fahrenheit") return celsius * 9.0 / 5.0 + 32;
            if (to == "Kelvin") return celsius + 273.15;
            throw new Exception("Unknown temperature unit");
        }

        private static double ConvertDensity(double value, string from, string to)
        {
            // First convert to kg/m³ (base unit)
            double kgPerM3;
            
            if (from == "API gravity")
            {
                // API to kg/m³: ρ = 141.5 / (API + 131.5) * 999.0
                kgPerM3 = 141.5 / (value + 131.5) * 999.0;
            }
            else if (from == "specific gravity")
            {
                // SG to kg/m³: ρ = SG * 999.0
                kgPerM3 = value * 999.0; // 999.0 is approx density of water at 15°C in kg/m³
            }
            else
            {
                // Regular conversion for other units
                kgPerM3 = from == "kilogram per cubic meter" ? value : 
                         value * GetDensityFactor(from);
            }
            
            // Now convert from kg/m³ to target unit
            if (to == "API gravity")
            {
                // kg/m³ to API: API = (141.5 / (ρ/999.0)) - 131.5
                return (141.5 / (kgPerM3/999.0)) - 131.5;
            }
            else if (to == "specific gravity")
            {
                // kg/m³ to SG: SG = ρ / 999.0
                return kgPerM3 / 999.0;
            }
            else
            {
                // Regular conversion for other units
                return to == "kilogram per cubic meter" ? kgPerM3 : 
                       kgPerM3 / GetDensityFactor(to);
            }
        }

        private static double GetDensityFactor(string unit)
        {
            switch(unit)
            {
                case "gram per cubic centimeter": return 1000.0;
                case "pound per cubic foot": return 16.0185;
                case "pound per gallon": return 119.8264;
                case "pound per barrel": return 0.158987 * 119.8264;
                default: throw new Exception($"Unknown density unit: {unit}");
            }
        }
    }

    public class UnitCategory
    {
        public string Name { get; }
        public Dictionary<string, double> Units { get; }
        public Dictionary<string, string> UnitSymbols { get; } // Add symbols dictionary
        
        public UnitCategory(string name, Dictionary<string, double> units, Dictionary<string, string> unitSymbols)
        {
            Name = name;
            Units = units;
            UnitSymbols = unitSymbols;
        }
    }

    public class UnitItem
    {
        public string Name { get; }
        public string Symbol { get; }
        
        public UnitItem(string name, string symbol)
        {
            Name = name;
            Symbol = symbol;
        }
        
        public override string ToString()
        {
            return $"{Name} [{Symbol}]";
        }
    }
}
