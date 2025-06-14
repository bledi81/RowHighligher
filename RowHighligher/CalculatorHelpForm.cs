﻿using System;
using System.Windows.Forms;
using System.Drawing;

namespace RowHighligher
{
    public partial class CalculatorHelpForm : Form
    {
        private TabControl tabControl;

        public CalculatorHelpForm()
        {
            InitializeComponents();
            this.TopMost = true; // Make help stay on top
        }

        private void InitializeComponents()
        {
            this.Text = "Scientific Calculator Help";
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

            // Settings Tab (Updated)
            TabPage settingsTab = CreateHelpTab("Settings", @"Calculator Settings

General Settings:
• Access settings via the Settings button
• Settings window provides configuration options
• Settings persist across Excel sessions

Decimal Places:
• Range: 0-10 decimal places
• Default: 4 decimal places
• Affects all calculations and displays
• Changes take effect immediately
• Settings persist during calculator session

Complex Number Insertion:
• Option: ""Add Complex result as formula""
• When checked: Complex results are inserted as Excel formulas
  Example: √(-9) becomes =COMPLEX(0, SQRT(9)
  Allowing Excel to perform further calculations with complex numbers
• When unchecked: Complex results are inserted as plain text
  Example: √(-9) becomes ""3i""


Examples with different decimal places:
• 0 places: 1/3 = 0
• 2 places: 1/3 = 0.33
• 4 places: 1/3 = 0.3333
• 6 places: 1/3 = 0.333333

Tips:
• Adjust decimal places based on needed precision
• Higher precision available for scientific calculations
• Settings window stays on top for easy access
• Click Save to apply changes
• Changes affect both display and calculations
• Complex numbers are formatted according to decimal places setting");

            // Basic Operations Tab (Updated)
            TabPage basicTab = CreateHelpTab("Basic Operations", @"Basic Operations

Display System:
• Blue display: Expression (input)
• Green display: Calculation result

Numbers and Basic Math:
• Numbers (0-9): Click buttons or type directly
• Decimal point (.) or (,): For decimal numbers
• Basic operators:
  + Addition       Example: 2 + 2 = 4
  - Subtraction    Example: 5 - 3 = 2
  × Multiplication Example: 4 × 3 = 12
  ÷ Division       Example: 10 ÷ 2 = 5

Input Methods:
• Use keyboard for direct entry
• Use on-screen buttons
• Get values from Excel with Get button

Memory Functions:
• MC: Memory Clear - Erases stored value
• MR: Memory Recall - Shows stored value
• M+: Memory Add - Adds display to memory
• M-: Memory Subtract - Subtracts from memory

Clear Functions:
• CE: Clear Entry - Clears current entry only
• ←: Backspace - Deletes last character
• LastAns: Recalls last calculation result

Excel Integration:
• Get: Gets value from current Excel cell
• Insert (Ctrl+Enter): Sends result to Excel
• Detachable window mode
• Always-on-top option when detached");

            // Scientific Functions Tab (Updated)
            TabPage scientificTab = CreateHelpTab("Scientific Functions", @"Scientific Functions

Constants:
• π (pi): Mathematical constant pi
  Example: π = 3.1416 (with 4 decimals)
  Type 'pi' or use π button

• e: Mathematical constant e
  Example: e = 2.7183 (with 4 decimals)
  Type 'e' or use e button

Mathematical Functions:
• sqrt(x): Square root
  Example: sqrt(16) = 4.0000
  Example: sqrt(2) ≈ 1.4142
  Note: sqrt(-1) returns complex number i

• x^y: Power function
  Example: 2^3 = 8.0000
  Type '^' or use x^y button

• 1/x: Reciprocal
  Example: 1/2 = 0.5000

Trigonometric Functions:
• sin(x): Sine function
• cos(x): Cosine function
• tan(x): Tangent function

Angle Modes:
• RAD: Radians mode (default)
• DEG: Degrees mode
• → RAD: Convert degrees to radians
• → DEG: Convert radians to degrees
  Example: 90° → RAD = π/2
  Example: π → DEG = 180°

Logarithmic Functions:
• log(x): Base-10 logarithm
  Example: log(100) = 2.0000
  Note: x must be positive

• ln(x): Natural logarithm
  Example: ln(e) = 1.0000
  Note: x must be positive");

            // Expression Mode Tab (Updated)
            TabPage expressionTab = CreateHelpTab("Expression Mode", @"Expression Mode & Syntax

Color-Coded Parentheses:
• Matching parentheses are colored the same
• Unmatched parentheses appear in red
• Visual indicator of expression structure

Parentheses Usage:
• Use ( ) for grouping operations
• Automatic parentheses balancing
• Visual bracket counting
• Example: (2 + 3) × 4 = 20
• Example: 2 × (3 + 4) = 14

Function Combinations:
• Functions can be nested
• Multiple operations in one expression
• Proper operator precedence
Examples:
• sqrt(sin(π/2) + 16) = 5
• log(ln(10)) ≈ 0.834
• 2 + sqrt(16) × 3 = 14

Order of Operations:
1. Parentheses ( )
2. Functions (sqrt, sin, cos, etc.)
3. Powers (^)
4. Multiplication and Division (×, ÷)
5. Addition and Subtraction (+, -)

Complex Number Support:
• Automatically handles complex numbers when needed
  Example: sqrt(-4) = 0 + 2i
• Properly calculates with complex intermediates
• Formats complex results with the correct precision

Expression Building Features:
• Real-time function name recognition
• Auto-completion for functions
• Intelligent operator spacing
• Automatic error checking
• Proper decimal handling");

            // Input Methods Tab (Updated)
            TabPage inputTab = CreateHelpTab("Input Methods", @"Input Methods & Shortcuts

Keyboard Mode:
• Blue display indicates keyboard mode
• Direct function typing with auto-completion:
  - Type 'sqrt' for square root
  - Type 'sin' for sine
  - Type 'cos' for cosine
  - Type 'tan' for tangent
  - Type 'log' for logarithm (base 10)
  - Type 'ln' for natural logarithm
  - Type 'pi' for π
  - Type 'e' for e

Special Keys:
• F1: Open this help window
• Ctrl+Enter: Insert result to Excel
• ESC: Clear calculator
• Backspace: Delete last character
• Enter or =: Calculate result

Mouse Input:
• Click buttons for numbers
• Click operators for operations
• Click functions for scientific operations
• Click parentheses for grouping
• Click LastAns for previous result
• Click Settings to configure

Excel Integration:
• Get: Import value from current Excel cell
• Insert: Send result to current Excel cell
• Insert as formula: Option in Settings menu
  - For complex numbers: =COMPLEX(real, imaginary)
  - For regular numbers: Value inserted directly
• Normal vs. formula mode determined in Settings
• Clipboard-friendly operations

Tips:
• Watch the color of the display
• Use parentheses for complex expressions
• Check bracket balance indicator
• Use CE to clear current entry
• Use LastAns to continue calculations
• Adjust decimal places in Settings");

            // Add tabs in order
            tabControl.TabPages.AddRange(new TabPage[] { 
                settingsTab,
                basicTab, 
                scientificTab, 
                expressionTab,
                inputTab
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