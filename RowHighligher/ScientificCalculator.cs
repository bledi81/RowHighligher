using System;
using System.Drawing;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;
using System.Data;
using System.Text;
using System.Collections.Generic; // Added for Stack
using System.Linq; // Added for LINQ operations
using System.Globalization; // Added for CultureInfo
using System.Numerics; // Ensure Complex is available for all classes

namespace RowHighligher
{
    public static class CalculatorConstants
    {
        public const double RAD_TO_DEG = 180.0 / Math.PI;
        public const double DEG_TO_RAD = Math.PI / 180.0;
    }

    public partial class ScientificCalculator : Form
    {
        private RichTextBox expressionDisplayTextBox; // Changed from TextBox to RichTextBox
        private TextBox resultDisplayTextBox;   // New TextBox for the result
        private Button insertButton;
        private Button clearButton;
        private Button helpButton;  // Add help button field
        private Button settingsButton; // Add settings button field
        private Button[] numberButtons = new Button[10]; // Initialize the array for digits 0-9
        private Button[] operatorButtons = new Button[6]; // Initialize array for operators +, -, *, /, =, .
        private Button[] scientificButtons;
        private double memory = 0;
        private bool isNewCalculation = true;
        private string lastOperation = "";
        private double lastValue = 0;
        private bool isInKeyboardMode = true;
        private bool isRadiansMode = true; // Track angle mode: true = radians, false = degrees

        // Add a field to track when an expression is being built
        private bool isExpressionMode = false;

        // Add this field near the top of your class with other fields
        private object lastAnswer = 0.0; // Changed to object to store double or Complex

        // Add this field to store characters as they're typed
        private string keyBuffer = "";

        // Add constants
        private const double PI = Math.PI;
        private const double E = Math.E;
        private const double RAD_TO_DEG = CalculatorConstants.RAD_TO_DEG;
        private const double DEG_TO_RAD = CalculatorConstants.DEG_TO_RAD;

        // Update the field to use the saved setting
        private int decimalPlaces = RowHighligher.Properties.Settings.Default.CalculatorDecimalPlaces;

        // Custom title bar fields
        private Panel customTitleBar;
        private Label titleLabel;
        private Button closeButton;
        private Button minimizeButton;
        private Point dragOffset;
        private bool dragging = false;

        // Add this near the top of your class with the other fields
        private TableLayoutPanel mainPanel;
        private TableLayoutPanel numberPanel;

        public ScientificCalculator()
        {
            InitializeComponents();

            // Set form properties
            this.FormBorderStyle = FormBorderStyle.None; // Use custom title bar
            this.Text = "Scientific Calculator";
            this.MinimumSize = new Size(350, 600);  // Increased minimum width & height
            this.Size = new Size(350, 600);         // Set larger default size
            this.MaximizeBox = false;
            this.MinimizeBox = true;
            this.TopMost = RowHighligher.Properties.Settings.Default.IsCalculatorDetached;
            this.StartPosition = FormStartPosition.CenterScreen;
            this.KeyPreview = true;

            // Add keyboard event handlers
            this.KeyDown += ScientificCalculator_KeyDown;
            this.KeyPress += ScientificCalculator_KeyPress;
            
            // Update event handlers to point to the new expressionDisplayTextBox
            this.expressionDisplayTextBox.GotFocus += DisplayTextBox_GotFocus;
            this.expressionDisplayTextBox.LostFocus += DisplayTextBox_LostFocus;
            this.expressionDisplayTextBox.TextChanged += ExpressionDisplayTextBox_TextChanged; // Add TextChanged event
        }

        private void ExpressionDisplayTextBox_TextChanged(object sender, EventArgs e)
        {
            HighlightParentheses();
        }

        // Add this to enhance the display's readability when showing expressions
        private void FormatDisplayText()
        {
            string text = expressionDisplayTextBox.Text;
            // Replace sqrt with √
            text = text.Replace("sqrt", "√");
            // Apply operator formatting regardless of mode
            text = text.Replace("*", " × ")
                       .Replace("/", " ÷ ");
            // Only format +/- if not in expression mode or no parentheses
            if (!isExpressionMode || (!text.Contains("(") && !text.Contains(")")))
            {
                text = text.Replace("+", " + ")
                          .Replace("-", " - ");
            }
            // Clean up spaces
            while (text.Contains("  "))
            {
                text = text.Replace("  ", " ");
            }
            expressionDisplayTextBox.Text = text.Trim();
            // Always move cursor to end after formatting
            SetCursorToEnd();
        }

        private void InitializeComponents()
        {
            // Custom title bar
            customTitleBar = new Panel
            {
                Dock = DockStyle.Fill,
                Height = 36,
                BackColor = Color.LightSkyBlue
            };
            titleLabel = new Label
            {
                Text = "Scientific Calculator",
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
                BackColor = Color.LightSkyBlue,
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
                BackColor = Color.LightSkyBlue,
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

            // Main layout panel
            mainPanel = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                RowCount = 7,         // Increased to 7 rows to accommodate separate result display
                ColumnCount = 1,
                Padding = new Padding(10)
            };

            // Update row styles for better proportions
            mainPanel.RowStyles.Add(new RowStyle(SizeType.Percent, 20));  // Expression Display (increased height)
            mainPanel.RowStyles.Add(new RowStyle(SizeType.Percent, 8));  // Result Display
            mainPanel.RowStyles.Add(new RowStyle(SizeType.Percent, 9));  // Buttons panel
            mainPanel.RowStyles.Add(new RowStyle(SizeType.Percent, 9));  // Parentheses panel
            mainPanel.RowStyles.Add(new RowStyle(SizeType.Percent, 21));  // Scientific panel
            mainPanel.RowStyles.Add(new RowStyle(SizeType.Percent, 26));  // Number panel (decreased slightly)
            mainPanel.RowStyles.Add(new RowStyle(SizeType.Percent, 7));   // Bottom panel (decreased slightly)

            // Expression display textbox
            expressionDisplayTextBox = new RichTextBox // Changed from TextBox
            {
                Dock = DockStyle.Fill,
                ReadOnly = true, // Keep ReadOnly for now, input via buttons/KeyPress
                Font = new Font("Consolas", 12, FontStyle.Regular),
                // TextAlign = HorizontalAlignment.Left, // Not applicable to RichTextBox directly
                Multiline = true,
                WordWrap = true, // WordWrap is true by default for RichTextBox if Multiline is true
                ScrollBars = RichTextBoxScrollBars.Vertical, // RichTextBox uses different enum
                Text = "",
                BackColor = Color.LightBlue // Set background color
            };

            mainPanel.Controls.Add(expressionDisplayTextBox, 0, 0);

            // Result display textbox
            resultDisplayTextBox = new TextBox
            {
                Dock = DockStyle.Fill,
                ReadOnly = true,
                Font = new Font("Consolas", 18, FontStyle.Bold), // Larger font for result
                TextAlign = HorizontalAlignment.Right,
                Text = "0", // Initial result display
                BackColor = Color.LightGreen // Set background color
            };
            mainPanel.Controls.Add(resultDisplayTextBox, 0, 1);


            // Add panel for display and decimal places control
            Panel displayPanel = new Panel // This panel seems redundant now, can be removed or re-evaluated.
            {                           // For now, let's assume it's not used as mainPanel directly holds textboxes.
                Dock = DockStyle.Fill,
                Padding = new Padding(0)
            };
            // displayPanel.Controls.Add(expressionDisplayTextBox); // This was the old setup

            // mainPanel.Controls.Add(displayPanel, 0, 0); // This was the old setup

            // Insert button in a panel with Clear and decimal places
            TableLayoutPanel buttonsPanel = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 5,  // Increase column count to add logo
                RowCount = 1
            };
            buttonsPanel.RowStyles.Add(new RowStyle(SizeType.Percent, 100f)); // Ensures the single row uses all available height
            for (int i = 0; i < 6; i++) // Update column count
            {
                buttonsPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, i == 4 ? 28 : 18)); // Logo column bigger
            }

            // Add Get button in the previous location of Help button
            Button getButton = new Button
            {
                Text = "Get",
                Dock = DockStyle.Fill,
                Font = new Font("Segoe UI", 8f, FontStyle.Regular),
                Margin = new Padding(1)
            };
            getButton.Click += GetButton_Click;
            buttonsPanel.Controls.Add(getButton, 0, 0);

            insertButton = new Button
            {
                Text = "Insert Ctrl+\u23CE",
                Dock = DockStyle.Fill,
                Font = new Font("Segoe UI", 8f, FontStyle.Regular),
                Margin = new Padding(1)
            };
            insertButton.Click += InsertButton_Click;
            buttonsPanel.Controls.Add(insertButton, 1, 0);

            clearButton = new Button
            {
                Text = "LastAns",
                Dock = DockStyle.Fill,
                Font = new Font("Segoe UI", 8f, FontStyle.Regular),
                Margin = new Padding(1)
            };
            clearButton.Click += LastAnsButton_Click;
            buttonsPanel.Controls.Add(clearButton, 2, 0);

            // Create a PictureBox for the logo
            PictureBox logoBox = new PictureBox
            {
                Dock = DockStyle.Fill,
                SizeMode = PictureBoxSizeMode.Zoom,
                Margin = new Padding(1),
                Image = RowHighligher.Properties.Resources.Bankers_Logo_Albania // Add your PNG image to resources
            };

            // Add logo after settings button
            buttonsPanel.Controls.Add(logoBox, 4, 0);

            // Add settings button
            settingsButton = new Button
            {
                Text = "Settings",
                Dock = DockStyle.Fill,
                Font = new Font("Segoe UI", 8f, FontStyle.Regular),
                Margin = new Padding(1)
            };
            settingsButton.Click += SettingsButton_Click;
            buttonsPanel.Controls.Add(settingsButton, 3, 0);

            // Adjust the row index for subsequent panels due to the new resultDisplayTextBox
            mainPanel.Controls.Add(buttonsPanel, 0, 2); // Moved from 0,1 to 0,2

            // Parentheses panel
            TableLayoutPanel parenthesesPanel = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 5, // Changed from 4 to 5 to accommodate Help button
                RowCount = 1
            };
            parenthesesPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 20));
            parenthesesPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 20));
            parenthesesPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 20));
            parenthesesPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 20));
            parenthesesPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 20));

            Button leftBracketButton = new Button
            {
                Text = "(",
                Dock = DockStyle.Fill,
                Tag = "("
            };
            leftBracketButton.Click += LeftBracketButton_Click;
            parenthesesPanel.Controls.Add(leftBracketButton, 0, 0);

            Button rightBracketButton = new Button
            {
                Text = ")",
                Dock = DockStyle.Fill,
                Tag = ")"
            };
            rightBracketButton.Click += RightBracketButton_Click;
            parenthesesPanel.Controls.Add(rightBracketButton, 1, 0);

            // Add backspace button
            Button backspaceButton = new Button
            {
                Text = "←",  // Unicode backspace arrow
                Dock = DockStyle.Fill,
                Tag = "backspace"
            };
            backspaceButton.Click += BackspaceButton_Click;
            parenthesesPanel.Controls.Add(backspaceButton, 2, 0);

            // Additional button (maybe "CE" - Clear Entry)
            Button clearEntryButton = new Button { Text = "CE", Dock = DockStyle.Fill, Tag = "clearEntry" };
            clearEntryButton.Click += ClearEntryButton_Click;
            parenthesesPanel.Controls.Add(clearEntryButton, 3, 0);
            
            // Move help button to the right of CE button
            helpButton = new Button
            {
                Text = "Help (F1)",
                Dock = DockStyle.Fill,
                Font = new Font("Segoe UI", 8f, FontStyle.Regular),
                Margin = new Padding(1)
            };
            helpButton.Click += HelpButton_Click;
            parenthesesPanel.Controls.Add(helpButton, 4, 0);

            // Add the parentheses panel to the main panel
            mainPanel.Controls.Add(parenthesesPanel, 0, 3); // Moved from 0,2 to 0,3

            // Create a panel to hold the scientific panel and provide the border
            Panel scientificBorderPanel = new Panel
            {
                Dock = DockStyle.Fill,
                BorderStyle = BorderStyle.FixedSingle,
                BackColor = Color.LightSteelBlue,
                Padding = new Padding(3)  // Increased padding for better appearance
            };

            // Scientific buttons panel
            TableLayoutPanel scientificPanel = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 4,  // Changed to 4 columns
                RowCount = 3,     // Changed to 3 rows
                BackColor = SystemColors.Control
            };

            // Set equal width for columns
            for (int i = 0; i < 4; i++)
            {
                scientificPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 25f));
            }

            // Set height for rows (making the constant row slightly smaller)
            scientificPanel.RowStyles.Add(new RowStyle(SizeType.Percent, 35f));  // First row
            scientificPanel.RowStyles.Add(new RowStyle(SizeType.Percent, 35f));  // Second row
            scientificPanel.RowStyles.Add(new RowStyle(SizeType.Percent, 30f));  // Constants row

            scientificButtons = new Button[12];  // Increased from 8 to 12
            string[] scientificOps = {
                "sqrt", "sin", "cos", "tan",      // Row 1
                "x^y", "log", "ln", "1/x",        // Row 2
                "π", "e", "RAD", "DEG"            // Row 3 (constants)
            };

            // Add buttons in 3x4 grid with improved styling
            for (int i = 0; i < scientificButtons.Length; i++)
            {
                int row = i / 4;    // 4 buttons per row
                int col = i % 4;    // Columns 0-3 in each row

                string buttonText = scientificOps[i];
                string buttonTag = scientificOps[i];

                // Special handling for RAD and DEG buttons
                if (buttonText == "RAD")
                {
                    buttonText = "→ RAD";
                    buttonTag = "deg_to_rad";
                }
                else if (buttonText == "DEG")
                {
                    buttonText = "→ DEG";
                    buttonTag = "rad_to_deg";
                }
                // Use proper symbol for sqrt
                else if (buttonText == "sqrt")
                {
                    buttonText = "√";
                    buttonTag = "sqrt";
                }
                // Use proper symbol for 1/x (optional, keep as is)

                scientificButtons[i] = new Button
                {
                    Text = buttonText,
                    Tag = buttonTag,
                    Dock = DockStyle.Fill,
                    Margin = new Padding(2),
                    Font = new Font("Segoe UI", 9.5f, FontStyle.Regular)
                };
                scientificButtons[i].Click += ScientificButton_Click;
                scientificPanel.Controls.Add(scientificButtons[i], col, row);
            }

            scientificBorderPanel.Controls.Add(scientificPanel);
            mainPanel.Controls.Add(scientificBorderPanel, 0, 4);  // Moved from 0,3 to 0,4

            // Set initial button colors for RAD/DEG mode
            UpdateAngleModeButtonsDisplay();

            // Number buttons panel for rows 3, 4, 5
            numberPanel = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 4,
                RowCount = 4,
                // Rest of the initialization remains unchanged
                RowStyles = {
                    new RowStyle(SizeType.Percent, 25),
                    new RowStyle(SizeType.Percent, 25),
                    new RowStyle(SizeType.Percent, 25),
                    new RowStyle(SizeType.Percent, 25)
                }
            };
            numberPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 25));
            numberPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 25));
            numberPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 25));
            numberPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 25));

            // Memory buttons and additional operators
            Button[] memButtons = new Button[4];
            string[] memOps = { "MC", "MR", "M+", "M-" };

            for (int i = 0; i < memButtons.Length; i++)
            {
                memButtons[i] = new Button
                {
                    Text = memOps[i],
                    Dock = DockStyle.Fill,
                    Tag = memOps[i]
                };
                memButtons[i].Click += MemoryButton_Click;
                numberPanel.Controls.Add(memButtons[i], i, 0);
            }

            // Number buttons 7-9 and ÷
            for (int i = 7; i <= 9; i++)
            {
                int col = i - 7;
                numberButtons[i] = new Button
                {
                    Text = i.ToString(),
                    Dock = DockStyle.Fill,
                    Tag = i.ToString()
                };
                numberButtons[i].Click += NumberButton_Click;
                numberPanel.Controls.Add(numberButtons[i], col, 1);
            }

            operatorButtons[0] = new Button { Text = "÷", Dock = DockStyle.Fill, Tag = "/", BackColor = Color.LightGoldenrodYellow };
            operatorButtons[0].Click += OperatorButton_Click;
            numberPanel.Controls.Add(operatorButtons[0], 3, 1);

            // Number buttons 4-6 and ×
            for (int i = 4; i <= 6; i++)
            {
                int col = i - 4;
                numberButtons[i] = new Button
                {
                    Text = i.ToString(),
                    Dock = DockStyle.Fill,
                    Tag = i.ToString()
                };
                numberButtons[i].Click += NumberButton_Click;
                numberPanel.Controls.Add(numberButtons[i], col, 2);
            }

            operatorButtons[1] = new Button { Text = "×", Dock = DockStyle.Fill, Tag = "*", BackColor = Color.LightGoldenrodYellow };
            operatorButtons[1].Click += OperatorButton_Click;
            numberPanel.Controls.Add(operatorButtons[1], 3, 2);

            // Number buttons 1-3 and -
            for (int i = 1; i <= 3; i++)
            {
                int col = i - 1;
                numberButtons[i] = new Button
                {
                    Text = i.ToString(),
                    Dock = DockStyle.Fill,
                    Tag = i.ToString()
                };
                numberButtons[i].Click += NumberButton_Click;
                numberPanel.Controls.Add(numberButtons[i], col, 3);
            }

            operatorButtons[2] = new Button { Text = "-", Dock = DockStyle.Fill, Tag = "-", BackColor = Color.LightGoldenrodYellow };
            operatorButtons[2].Click += OperatorButton_Click;
            numberPanel.Controls.Add(operatorButtons[2], 3, 3);

            mainPanel.Controls.Add(numberPanel, 0, 5);  // Moved from 0,4 to 0,5

            // Bottom panel for 0, ., +/-, =
            TableLayoutPanel bottomPanel = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 4,
                RowCount = 1
            };
            bottomPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 25));
            bottomPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 25));
            bottomPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 25));
            bottomPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 25));

            numberButtons[0] = new Button { Text = "0", Dock = DockStyle.Fill, Tag = "0" };
            numberButtons[0].Click += NumberButton_Click;
            bottomPanel.Controls.Add(numberButtons[0], 0, 0);

            operatorButtons[3] = new Button { Text = ".", Dock = DockStyle.Fill, Tag = "." };
            operatorButtons[3].Click += DecimalPoint_Click;
            bottomPanel.Controls.Add(operatorButtons[3], 1, 0);

            operatorButtons[4] = new Button { Text = "=", Dock = DockStyle.Fill, Tag = "=", BackColor = Color.Orange };
            operatorButtons[4].Click += EqualsButton_Click;
            bottomPanel.Controls.Add(operatorButtons[4], 2, 0);

            operatorButtons[5] = new Button { Text = "+", Dock = DockStyle.Fill, Tag = "+", BackColor = Color.LightGoldenrodYellow };
            operatorButtons[5].Click += OperatorButton_Click;
            bottomPanel.Controls.Add(operatorButtons[5], 3, 0);

            mainPanel.Controls.Add(bottomPanel, 0, 6); // Moved from 0,5 to 0,6

            // Root layout: 2 rows, title bar and calculator UI
            var rootPanel = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                RowCount = 2,
                ColumnCount = 1
            };
            rootPanel.RowStyles.Add(new RowStyle(SizeType.Absolute, 36)); // Title bar height
            rootPanel.RowStyles.Add(new RowStyle(SizeType.Percent, 100)); // Calculator UI
            rootPanel.Controls.Add(customTitleBar, 0, 0);
            rootPanel.Controls.Add(mainPanel, 0, 1);
            this.Controls.Clear();
            this.Controls.Add(rootPanel);
        }

        private void HighlightParentheses()
        {
            if (expressionDisplayTextBox.Text.Length == 0)
                return;

            // Save current selection
            int selectionStart = expressionDisplayTextBox.SelectionStart;
            int selectionLength = expressionDisplayTextBox.SelectionLength;

            // Reset all text to default color (e.g., black)
            expressionDisplayTextBox.SelectAll();
            expressionDisplayTextBox.SelectionColor = Color.Black; // Or your default text color
            expressionDisplayTextBox.DeselectAll();

            Stack<int> openParenthesesIndices = new Stack<int>();
            // Define a few colors for matching pairs - can be extended
            Color[] pairColors = { Color.Blue, Color.Green, Color.Purple, Color.OrangeRed }; 
            int colorIndex = 0;

            for (int i = 0; i < expressionDisplayTextBox.Text.Length; i++)
            {
                char c = expressionDisplayTextBox.Text[i];
                if (c == '(')
                {
                    openParenthesesIndices.Push(i);
                }
                else if (c == ')')
                {
                    if (openParenthesesIndices.Count > 0)
                    {
                        int openIndex = openParenthesesIndices.Pop();
                        Color currentColor = pairColors[colorIndex % pairColors.Length];
                        colorIndex++;

                        // Highlight opening parenthesis
                        expressionDisplayTextBox.Select(openIndex, 1);
                        expressionDisplayTextBox.SelectionColor = currentColor;

                        // Highlight closing parenthesis
                        expressionDisplayTextBox.Select(i, 1);
                        expressionDisplayTextBox.SelectionColor = currentColor;
                    }
                    else
                    {
                        // Unmatched closing parenthesis - color it red
                        expressionDisplayTextBox.Select(i, 1);
                        expressionDisplayTextBox.SelectionColor = Color.Red;
                    }
                }
            }

            // Any remaining open parentheses are unmatched - color them red
            while (openParenthesesIndices.Count > 0)
            {
                int unmatchedOpenIndex = openParenthesesIndices.Pop();
                expressionDisplayTextBox.Select(unmatchedOpenIndex, 1);
                expressionDisplayTextBox.SelectionColor = Color.Red;
            }

            // Restore original selection
            expressionDisplayTextBox.Select(selectionStart, selectionLength);
            // Ensure focus remains if it was there
            if(this.ActiveControl == expressionDisplayTextBox) expressionDisplayTextBox.Focus();
        }

        // Add a field to track bracket balance
        private int bracketCount = 0;

        private void LeftBracketButton_Click(object sender, EventArgs e)
        {
            AppendToDisplay("(");
            bracketCount++;
        }

        private void RightBracketButton_Click(object sender, EventArgs e)
        {
            // Only allow closing bracket if there are open brackets
            if (bracketCount > 0)
            {
                AppendToDisplay(")");
                bracketCount--;

                // If all brackets are closed, we're no longer in expression mode
                if (bracketCount == 0)
                {
                    isExpressionMode = false;
                }
            }
        }

        private void ScientificCalculator_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Control && e.KeyCode == Keys.Enter)
            {
                InsertButton_Click(sender, e);
                e.Handled = true;
            }
            else if (e.KeyCode == Keys.Escape) // Add this condition to handle ESC key
            {
                ClearButton_Click(sender, e);
                e.Handled = true;
            }
        }

        private void ScientificCalculator_KeyPress(object sender, KeyPressEventArgs e)
        {
            // Only process key presses if keyboard mode is enabled
            if (!isInKeyboardMode)
                return;

            // Handle letters for function names (sqrt, sin, cos)
            if (char.IsLetter(e.KeyChar))
            {
                // Add the letter to the buffer
                keyBuffer += e.KeyChar;
                
                // Display what the user is typing in real-time
                if (isNewCalculation || expressionDisplayTextBox.Text == "0" || string.IsNullOrEmpty(expressionDisplayTextBox.Text)) // Check for empty too
                {
                    expressionDisplayTextBox.Text = keyBuffer;
                }
                else
                {
                    expressionDisplayTextBox.Text += e.KeyChar; // Add just the current character
                }
                
                // Check if we've completed a function name
                if (keyBuffer.EndsWith("sqrt", StringComparison.OrdinalIgnoreCase) ||
                    keyBuffer.EndsWith("sin", StringComparison.OrdinalIgnoreCase) ||
                    keyBuffer.EndsWith("cos", StringComparison.OrdinalIgnoreCase) ||
                    keyBuffer.EndsWith("tan", StringComparison.OrdinalIgnoreCase) ||
                    keyBuffer.EndsWith("log", StringComparison.OrdinalIgnoreCase) ||
                    keyBuffer.EndsWith("ln", StringComparison.OrdinalIgnoreCase) ||
                    keyBuffer.EndsWith("pi", StringComparison.OrdinalIgnoreCase) ||  // Add pi
                    keyBuffer.Equals("e", StringComparison.OrdinalIgnoreCase))       // Add e
                {
                    // Get function name
                    string functionName = "";
                    int charsToCut = 0;
                    
                    if (keyBuffer.EndsWith("sqrt", StringComparison.OrdinalIgnoreCase))
                    {
                        functionName = "sqrt";
                        charsToCut = 4;
                    }
                    else if (keyBuffer.EndsWith("sin", StringComparison.OrdinalIgnoreCase))
                    {
                        functionName = "sin";
                        charsToCut = 3;
                    }
                    else if (keyBuffer.EndsWith("cos", StringComparison.OrdinalIgnoreCase))
                    {
                        functionName = "cos"; 
                        charsToCut = 3;
                    }
                    else if (keyBuffer.EndsWith("tan", StringComparison.OrdinalIgnoreCase))
                    {
                        functionName = "tan";
                        charsToCut = 3;
                    }
                    else if (keyBuffer.EndsWith("log", StringComparison.OrdinalIgnoreCase))
                    {
                        functionName = "log";
                        charsToCut = 3;
                    }
                    else if (keyBuffer.EndsWith("ln", StringComparison.OrdinalIgnoreCase))
                    {
                        functionName = "ln";
                        charsToCut = 2;
                    }
                    else if (keyBuffer.EndsWith("pi", StringComparison.OrdinalIgnoreCase))
                    {
                        // Special handling for pi
                        string currentTextInner = expressionDisplayTextBox.Text; // Renamed to 'currentTextInner'
                        expressionDisplayTextBox.Text = currentTextInner.Substring(0, currentTextInner.Length - 2) + "π";
                        keyBuffer = "";
                        isNewCalculation = false;
                        SetCursorToEnd();
                        e.Handled = true;
                        return;
                    }
                    else if (keyBuffer.Equals("e", StringComparison.OrdinalIgnoreCase))
                    {
                        // Special handling for e
                        string currentTextInner = expressionDisplayTextBox.Text; // Renamed to 'currentTextInner'
                        expressionDisplayTextBox.Text = currentTextInner.Substring(0, currentTextInner.Length - 1) + "e";
                        keyBuffer = "";
                        isNewCalculation = false;
                        SetCursorToEnd();
                        e.Handled = true;
                        return;
                    }
                    
                    // Remove the function characters and add the function with parenthesis
                    string currentTextOuter = expressionDisplayTextBox.Text;
                    if (functionName == "sqrt")
                    {
                        expressionDisplayTextBox.Text = currentTextOuter.Substring(0, currentTextOuter.Length - charsToCut) + "√(";
                    }
                    else
                    {
                        expressionDisplayTextBox.Text = currentTextOuter.Substring(0, currentTextOuter.Length - charsToCut) + functionName + "(";
                    }
                    
                    // Reset buffer since we've processed this function
                    keyBuffer = "";
                    
                    // Update state variables for expression mode
                    isExpressionMode = true;
                    bracketCount++;
                    isNewCalculation = false;
                    
                    SetCursorToEnd();
                }
                
                // Buffer should not grow too large
                if (keyBuffer.Length > 10)
                {
                    keyBuffer = keyBuffer.Substring(keyBuffer.Length - 10);
                }
                
                e.Handled = true;
                return;
            }
            
            // Clear the buffer when non-letter keys are pressed
            if (!char.IsLetter(e.KeyChar))
            {
                keyBuffer = "";
            }
            
            // Rest of your existing key handling code...
            // Add handling for parentheses
            if (e.KeyChar == '(')
            {
                AppendToDisplay("(");
                bracketCount++;
                e.Handled = true;
            }
            else if (e.KeyChar == ')')
            {
                if (bracketCount > 0)
                {
                    AppendToDisplay(")");
                    bracketCount--;
                }
                e.Handled = true;
            }

            else if (e.KeyChar == '\b') // Handle backspace
            {
                if (expressionDisplayTextBox.Text.Length > 0)
                {
                    // Check if we're deleting a parenthesis
                    char lastChar = expressionDisplayTextBox.Text[expressionDisplayTextBox.Text.Length - 1];
                    if (lastChar == '(')
                        bracketCount--;
                    else if (lastChar == ')')
                        bracketCount++;

                    expressionDisplayTextBox.Text = expressionDisplayTextBox.Text.Substring(0, expressionDisplayTextBox.Text.Length - 1);
                    if (expressionDisplayTextBox.Text.Length == 0)
                    {
                        // expressionDisplayTextBox.Text = "0"; // Reset to 0 if empty - better to leave it empty for expression
                        resultDisplayTextBox.Text = "0"; // Reset result display
                    }
                    FormatDisplayText(); // Call to format the display text for better readability
                }
                e.Handled = true;
            }

            // Handle number keys
            if (char.IsDigit(e.KeyChar))
            {
                AppendToDisplay(e.KeyChar.ToString());
                e.Handled = true;
            }
            // Handle operators
            else if (e.KeyChar == '+' || e.KeyChar == '-' || e.KeyChar == '*' || e.KeyChar == '/')
            {
                HandleOperation(e.KeyChar.ToString());
                e.Handled = true;
            }
            // Handle decimal point - use the improved HandleDecimalPoint method
            else if (e.KeyChar == '.' || e.KeyChar == ',')
            {
                HandleDecimalPoint();
                e.Handled = true;
            }
            // Handle equals
            else if (e.KeyChar == '=' || e.KeyChar == '\r')
            {
                CalculateResult();
                e.Handled = true;
            }
            // In your ScientificCalculator_KeyPress method
            else if (e.KeyChar == '^')
            {
                // Add the power operator
                expressionDisplayTextBox.Text += "^";
                isExpressionMode = true; // Treat power operations as expressions
                isNewCalculation = false;
                SetCursorToEnd();
                e.Handled = true;
            }

            // Always show math symbols in the display
            expressionDisplayTextBox.Text = expressionDisplayTextBox.Text
                .Replace("sqrt", "√")
                .Replace("*", "×")
                .Replace("/", "÷");
            SetCursorToEnd();
        }

        private void DisplayTextBox_GotFocus(object sender, EventArgs e)
        {
            isInKeyboardMode = true;
            // expressionDisplayTextBox.BackColor = Color.LightSkyBlue; // Keep LightGreen or choose another focus color
        }

        private void DisplayTextBox_LostFocus(object sender, EventArgs e)
        {
            isInKeyboardMode = false;
            // expressionDisplayTextBox.BackColor = SystemColors.Window; // Keep LightGreen
        }

        private void NumberButton_Click(object sender, EventArgs e)
        {
            Button button = (Button)sender;
            AppendToDisplay(button.Text);
        }

        private void AppendToDisplay(string value)
        {
            if (isNewCalculation)
            {
                expressionDisplayTextBox.Text = ""; // Clear expression display for new calculation
                resultDisplayTextBox.Text = "0";    // Reset result display
                // For starting a new calculation
                if (value == "sqrt" || value == "√" || value == "sin" || value == "cos" || value == "tan" || value == "log" || value == "ln")
                {
                    // For functions, show the function name followed by opening parenthesis
                    if (value == "sqrt" || value == "√")
                        expressionDisplayTextBox.Text = "√(";
                    else
                        expressionDisplayTextBox.Text = value + "(";
                    isExpressionMode = true;
                    bracketCount++; // Important: increment bracket count when adding function parenthesis
                }
                else if (value == "(")
                {
                    // For opening parenthesis
                    expressionDisplayTextBox.Text = value;
                    bracketCount++;
                    isExpressionMode = true; // Important: set expression mode when opening parenthesis
                }
                else if (value == "-")
                {
                    // For negative number
                    expressionDisplayTextBox.Text = value;
                }
                else if (char.IsDigit(value[0]) || value == ".")
                {
                    // For digits or decimal point
                    expressionDisplayTextBox.Text = value;
                }
                else
                {
                    // For operators, start with 0 then operator (or previous answer)
                    // expressionDisplayTextBox.Text = "0" + value; // Old behavior
                    expressionDisplayTextBox.Text = resultDisplayTextBox.Text + value; // Use previous result
                }
                isNewCalculation = false;
            }
            else
            {
                // For continuing an expression
                if (value == "sqrt" || value == "√" || value == "sin" || value == "cos" || value == "tan" || value == "log" || value == "ln")
                {
                    if (value == "sqrt" || value == "√")
                        expressionDisplayTextBox.Text += "√(";
                    else
                        expressionDisplayTextBox.Text += value + "(";
                    isExpressionMode = true;
                    bracketCount++; // Important: increment bracket count when adding function parenthesis
                }
                else if (string.IsNullOrEmpty(expressionDisplayTextBox.Text) && value != ".") // Was displayTextBox.Text == "0"
                {
                    expressionDisplayTextBox.Text = value;
                }
                else
                {
                    expressionDisplayTextBox.Text += value;
                }
            }

            // After appending the value, always check if we should be in expression mode
            // If there are any open brackets, we should be in expression mode
            isExpressionMode = bracketCount > 0 ||
                              expressionDisplayTextBox.Text.Contains("(") && !HasBalancedParentheses(expressionDisplayTextBox.Text);


            // Only format the display once, and only when appropriate
            // if (!isExpressionMode) // Formatting logic might need adjustment with RichTextBox
            // {
            //     // FormatDisplayText(); 
            // }
            // HighlightParentheses(); // Called by TextChanged event
            SetCursorToEnd();
        }

        private void DecimalPoint_Click(object sender, EventArgs e)
        {
            HandleDecimalPoint();
        }

        private void HandleDecimalPoint()
        {
            if (isNewCalculation)
            {
                expressionDisplayTextBox.Text = "0.";
                isNewCalculation = false;
            }
            else
            {
                // Check if we need to handle multiple terms
                string currentText = expressionDisplayTextBox.Text;
                
                // Find operators to identify terms
                char[] operators = new[] { '+', '-', '*', '/', '×', '÷', '^' };
                int lastOperatorIndex = -1;
                
                // Find the last operator in the expression to identify the current term
                foreach (char op in operators)
                {
                    int index = currentText.LastIndexOf(op);
                    if (index > lastOperatorIndex)
                        lastOperatorIndex = index;
                }
                
                // If we have an operator (working with multiple terms)
                if (lastOperatorIndex >= 0)
                {
                    // Extract the current term (everything after the last operator)
                    string currentTerm = currentText.Substring(lastOperatorIndex + 1).Trim();
                    
                    // Only add decimal point if this term doesn't already have one
                    if (!currentTerm.Contains("."))
                    {
                        // If the term is empty, add "0."
                        if (string.IsNullOrWhiteSpace(currentTerm))
                            expressionDisplayTextBox.Text += "0.";
                        else
                            expressionDisplayTextBox.Text += ".";
                    }
                }
                else
                {
                    // Working with just the first term
                    if (!currentText.Contains("."))
                        expressionDisplayTextBox.Text += ".";
                }
            }
            
            // Format and set cursor position
            // FormatDisplayText(); // Consider when to call this
            SetCursorToEnd();
        }

        private void OperatorButton_Click(object sender, EventArgs e)
        {
            Button button = (Button)sender;
            HandleOperation(button.Text);
        }

        private void HandleOperation(string operation)
        {
            // Always allow operators after a result or in expression mode
            if (operation == "-" && (isNewCalculation || string.IsNullOrEmpty(expressionDisplayTextBox.Text))) // was displayTextBox.Text == "0"
            {
                // Special case for negative numbers at start
                expressionDisplayTextBox.Text = operation;
                isNewCalculation = false;
            }
            else if (isNewCalculation)
            {
                // If starting fresh with non-negative operator, prepend last answer or 0
                expressionDisplayTextBox.Text = resultDisplayTextBox.Text + " " + operation + " ";
                lastOperation = operation;
                isNewCalculation = false;
            }
            else
            {
                // For continuing an expression or after result
                expressionDisplayTextBox.Text += " " + operation + " ";
                lastOperation = operation;
                isNewCalculation = false;
            }
            
            SetCursorToEnd();
        }

        private void EqualsButton_Click(object sender, EventArgs e)
        {
            CalculateResult();
            lastValue = lastAnswer is double ? (double)lastAnswer : 0; // Store the last calculated result in lastValue
        }

        private void CalculateResult()
        {
            try
            {
                if (isNewCalculation || string.IsNullOrEmpty(expressionDisplayTextBox.Text)) // was displayTextBox.Text == "0"
                {
                    return; // No calculation needed
                }

                if (!HasBalancedParentheses(expressionDisplayTextBox.Text))
                {
                    if (expressionDisplayTextBox.Text.Contains("(") || expressionDisplayTextBox.Text.Contains(")"))
                    {
                        MessageBox.Show("Expression has unbalanced parentheses. Please close all open brackets.",
                            "Syntax Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        bracketCount = CountOpenParentheses(expressionDisplayTextBox.Text);
                        return;
                    }
                }

                isExpressionMode = false;
                bracketCount = 0;

                // Replace all Unicode square root symbols with "sqrt"
                string evalExpr = expressionDisplayTextBox.Text.Replace("√", "sqrt");
                object resultObject = ExpressionEvaluator.Evaluate(evalExpr, decimalPlaces, isRadiansMode);
                lastAnswer = resultObject;

                if (resultObject is Complex complexResult)
                {
                    resultDisplayTextBox.Text = FormatComplex(complexResult, decimalPlaces);
                }
                else if (resultObject is double doubleResult)
                {
                    resultDisplayTextBox.Text = doubleResult.ToString("F" + decimalPlaces, CultureInfo.InvariantCulture);
                }
                else
                {
                    // Should not happen if Evaluate returns double or throws specific exception for Complex
                    resultDisplayTextBox.Text = "Error"; 
                }
                
                lastOperation = "";
                isNewCalculation = true; 
                SetCursorToEnd();
            }
            catch (ArgumentException ex) when (ex.Message == "NegativeSqrt")
            {
                if (ex.Data.Contains("ComplexResult") && ex.Data["ComplexResult"] is Complex complexVal)
                {
                    lastAnswer = complexVal;
                    resultDisplayTextBox.Text = FormatComplex(complexVal, decimalPlaces);
                    isNewCalculation = true;
                    SetCursorToEnd();
                }
                else
                {
                    ShowError("Error calculating square root.");
                }
            }
            catch (Exception ex)
            {
                ShowError("Error calculating result: " + ex.Message);
            }
        }
        
        private string FormatComplex(Complex c, int dp)
        {
            string realPart = c.Real.ToString("F" + dp, CultureInfo.InvariantCulture);
            string imagPartAbs = Math.Abs(c.Imaginary).ToString("F" + dp, CultureInfo.InvariantCulture);

            if (Math.Abs(c.Imaginary) < Math.Pow(10, -dp-1)) // Treat very small imaginary part as zero
            {
                return realPart;
            }
            if (Math.Abs(c.Real) < Math.Pow(10, -dp-1)) // Treat very small real part as zero
            {
                 return (c.Imaginary < 0 ? "-" : "") + imagPartAbs + "i";
            }
            return $"{realPart} {(c.Imaginary < 0 ? "-" : "+")} {imagPartAbs}i";
        }

        private void ShowError(string message)
        {
            resultDisplayTextBox.Text = "Error"; // Show error in result display
            expressionDisplayTextBox.Text = ""; // Clear expression display
            isNewCalculation = true;
            isExpressionMode = false;
            MessageBox.Show(message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }

        // Add this helper method to count open parentheses
        private int CountOpenParentheses(string expression)
        {
            int count = 0;
            foreach (char c in expression)
            {
                if (c == '(') count++;
                else if (c == ')') count--;
            }
            return Math.Max(0, count); // Ensure we don't return a negative value
        }

        private string PrepareExpressionForEvaluation(string expression)
        {
            // Replace symbolic constants with their values before evaluation
            expression = expression.Replace("π", PI.ToString(CultureInfo.InvariantCulture));
            expression = expression.Replace("e", E.ToString(CultureInfo.InvariantCulture));
            // Replace √ with sqrt for evaluation
            expression = expression.Replace("√", "sqrt");
            // Rest of the method remains the same...
            while (expression.Contains("^"))
            {
                int powerIndex = expression.IndexOf("^");
                
                // Find the base (number before ^)
                int baseStartIndex = powerIndex - 1;
                while (baseStartIndex >= 0 && 
                      (char.IsDigit(expression[baseStartIndex]) || 
                       expression[baseStartIndex] == '.' ||
                       expression[baseStartIndex] == ')'))
                {
                    baseStartIndex--;
                }
                baseStartIndex++; // Adjust to actual start
                
                // Find the exponent (number after ^)
                int exponentEndIndex = powerIndex + 1;
                while (exponentEndIndex < expression.Length && 
                      (char.IsDigit(expression[exponentEndIndex]) || 
                       expression[exponentEndIndex] == '.' || 
                       expression[exponentEndIndex] == '('))
                {
                    exponentEndIndex++;
                }
                
                // Extract the base and exponent
                string baseStr = expression.Substring(baseStartIndex, powerIndex - baseStartIndex);
                string exponentStr = expression.Substring(powerIndex + 1, exponentEndIndex - powerIndex - 1);
                
                // Evaluate the power using Math.Pow
                DataTable dt = new DataTable();
                double baseValue = Convert.ToDouble(dt.Compute(baseStr, ""), CultureInfo.InvariantCulture);
                double exponentValue = Convert.ToDouble(dt.Compute(exponentStr, ""), CultureInfo.InvariantCulture);
                double powerResult = Math.Pow(baseValue, exponentValue);
                
                // Replace the power expression with the result
                string replacement = "(" + powerResult.ToString($"F{decimalPlaces}", CultureInfo.InvariantCulture) + ")";
                expression = expression.Substring(0, baseStartIndex) + 
                             replacement + 
                             expression.Substring(exponentEndIndex);
            }

            expression = expression.Replace(" ", "")
                                  .Replace("×", "*")
                                  .Replace("÷", "/");

            try {
                // Process functions one at a time
                while (expression.Contains("sqrt(") || expression.Contains("sin(") || 
                       expression.Contains("cos(") || expression.Contains("tan(") ||
                       expression.Contains("log(") || expression.Contains("ln("))
                {
                    // Process sqrt functions
                    int sqrtIdx = expression.IndexOf("sqrt(");
                    if (sqrtIdx >= 0)
                    {
                        // Find the matching closing parenthesis
                        int openCount = 1;
                        int closeIdx = sqrtIdx + 5;
                        
                        while (openCount > 0 && closeIdx < expression.Length)
                        {
                            if (expression[closeIdx] == '(') openCount++;
                            if (expression[closeIdx] == ')') openCount--;
                            closeIdx++;
                        }
                        
                        // Make sure we found the closing parenthesis
                        if (openCount == 0)
                        {
                            // Extract inner expression and evaluate it
                            string innerExpr = expression.Substring(sqrtIdx + 5, closeIdx - sqrtIdx - 6);
                            
                            // If there are functions inside, recursively evaluate those first
                            if (innerExpr.Contains("sqrt(") || innerExpr.Contains("sin(") || innerExpr.Contains("cos("))
                            {
                                innerExpr = PrepareExpressionForEvaluation(innerExpr); // Recursive call
                            }
                            
                            // Calculate the sqrt of the inner expression
                            DataTable dt = new DataTable();
                            double innerValue = Convert.ToDouble(dt.Compute(innerExpr, ""), CultureInfo.InvariantCulture);
                            
                            if (innerValue < 0)
                            {
                                Complex complexResult = Complex.Sqrt(new Complex(innerValue, 0));
                                ArgumentException ex = new ArgumentException("NegativeSqrt");
                                ex.Data["ComplexResult"] = complexResult;
                                throw ex;
                            }
                            double sqrtValue = Math.Sqrt(innerValue);
                            
                            string replacement = "(" + sqrtValue.ToString($"F{decimalPlaces}", CultureInfo.InvariantCulture) + ")";
                            expression = expression.Substring(0, sqrtIdx) + replacement + expression.Substring(closeIdx);
                            
                            // Since we've modified the expression, we need to process it again
                            return PrepareExpressionForEvaluation(expression); // Return modified string
                        }
                    }
                    
                    // Process sin functions
                    int sinIdx = expression.IndexOf("sin(");
                    if (sinIdx >= 0)
                    {
                        // Find the matching closing parenthesis
                        int openCount = 1;
                        int closeIdx = sinIdx + 4;  // Start after "sin("
                        
                        while (openCount > 0 && closeIdx < expression.Length)
                        {
                            if (expression[closeIdx] == '(') openCount++;
                            if (expression[closeIdx] == ')') openCount--;
                            closeIdx++;
                        }
                        
                        if (openCount == 0) // Found the matching closing parenthesis
                        {
                            // Extract and evaluate the inner expression
                            string innerExpr = expression.Substring(sinIdx + 4, closeIdx - sinIdx - 5);
                            
                            // Handle nested functions
                            if (innerExpr.Contains("sqrt(") || innerExpr.Contains("sin(") || innerExpr.Contains("cos("))
                            {
                                innerExpr = PrepareExpressionForEvaluation(innerExpr); // Recursive call
                            }
                            
                            // Calculate sin of the inner expression
                            DataTable dt = new DataTable();
                            double innerValue = Convert.ToDouble(dt.Compute(innerExpr, ""), CultureInfo.InvariantCulture);
                            
                            // Apply degrees to radians conversion if needed (assuming isRadiansMode is accessible or passed)
                            // For simplicity, this example doesn't pass isRadiansMode here. 
                            // This should be addressed if PrepareExpressionForEvaluation is used directly.
                            // if (!isRadiansMode) innerValue *= DEG_TO_RAD; 
                                
                            double sinValue = Math.Sin(innerValue); // Requires isRadiansMode
                            
                            // Replace with result (in parentheses)
                            string replacement = "(" + sinValue.ToString($"F{decimalPlaces}", CultureInfo.InvariantCulture) + ")";
                            expression = expression.Substring(0, sinIdx) + replacement + expression.Substring(closeIdx);
                            
                            // Process the updated expression
                            return PrepareExpressionForEvaluation(expression); // Return modified string
                        }
                    }
                    
                    // Process cos functions - same pattern as sin
                    int cosIdx = expression.IndexOf("cos(");
                    if (cosIdx >= 0)
                    {
                        int openCount = 1;
                        int closeIdx = cosIdx + 4;  // Start after "cos("
                        
                        while (openCount > 0 && closeIdx < expression.Length)
                        {
                            if (expression[closeIdx] == '(') openCount++;
                            if (expression[closeIdx] == ')') openCount--;
                            closeIdx++;
                        }
                        
                        if (openCount == 0)
                        {
                            string innerExpr = expression.Substring(cosIdx + 4, closeIdx - cosIdx - 5);
                            
                            if (innerExpr.Contains("sqrt(") || innerExpr.Contains("sin(") || innerExpr.Contains("cos("))
                            {
                                innerExpr = PrepareExpressionForEvaluation(innerExpr); // Recursive call
                            }
                            
                            DataTable dt = new DataTable();
                            double innerValue = Convert.ToDouble(dt.Compute(innerExpr, ""), CultureInfo.InvariantCulture);
                            // if (!isRadiansMode) innerValue *= CalculatorConstants.DEG_TO_RAD; // Requires isRadiansMode
                            double cosValue = Math.Cos(innerValue); // Requires isRadiansMode logic
                            
                            string replacement = "(" + cosValue.ToString($"F{decimalPlaces}", CultureInfo.InvariantCulture) + ")";
                            expression = expression.Substring(0, cosIdx) + replacement + expression.Substring(closeIdx);
                            
                            return PrepareExpressionForEvaluation(expression); // Return modified string
                        }
                    }

                    // Process tan functions
                    int tanIdx = expression.IndexOf("tan(");
                    if (tanIdx >= 0)
                    {
                        int openCount = 1;
                        int closeIdx = tanIdx + 4;  // Start after "tan("
                        
                        while (openCount > 0 && closeIdx < expression.Length)
                        {
                            if (expression[closeIdx] == '(') openCount++;
                            if (expression[closeIdx] == ')') openCount--;
                            closeIdx++;
                        }
                        
                        if (openCount == 0)
                        {
                            string innerExpr = expression.Substring(tanIdx + 4, closeIdx - tanIdx - 5);
                            
                            if (innerExpr.Contains("sqrt(") || innerExpr.Contains("sin(") || 
                                innerExpr.Contains("cos(") || innerExpr.Contains("tan(") ||
                                innerExpr.Contains("log(") || innerExpr.Contains("ln("))
                            {
                                innerExpr = PrepareExpressionForEvaluation(innerExpr); // Recursive call
                            }
                            
                            DataTable dt = new DataTable();
                            double innerValue = Convert.ToDouble(dt.Compute(innerExpr, ""), CultureInfo.InvariantCulture);
                            // if (!isRadiansMode) innerValue *= CalculatorConstants.DEG_TO_RAD; // Requires isRadiansMode
                            double tanValue = Math.Tan(innerValue); // Requires isRadiansMode logic
                            
                            string replacement = "(" + tanValue.ToString($"F{decimalPlaces}", CultureInfo.InvariantCulture) + ")";
                            expression = expression.Substring(0, tanIdx) + replacement + expression.Substring(closeIdx);
                            
                            return PrepareExpressionForEvaluation(expression); // Return modified string
                        }
                    }

                    // Process log functions (base 10)
                    int logIdx = expression.IndexOf("log(");
                    if (logIdx >= 0)
                    {
                        int openCount = 1;
                        int closeIdx = logIdx + 4;  // Start after "log("
                        
                        while (openCount > 0 && closeIdx < expression.Length)
                        {
                            if (expression[closeIdx] == '(') openCount++;
                            if (expression[closeIdx] == ')') openCount--;
                            closeIdx++;
                        }
                        
                        if (openCount == 0)
                        {
                            string innerExpr = expression.Substring(logIdx + 4, closeIdx - logIdx - 5);
                            
                            if (innerExpr.Contains("sqrt(") || innerExpr.Contains("sin(") || 
                                innerExpr.Contains("cos(") || innerExpr.Contains("tan(") ||
                                innerExpr.Contains("log(") || innerExpr.Contains("ln("))
                            {
                                innerExpr = PrepareExpressionForEvaluation(innerExpr); // Recursive call
                            }
                            
                            DataTable dt = new DataTable();
                            double innerValue = Convert.ToDouble(dt.Compute(innerExpr, ""), CultureInfo.InvariantCulture);
                            if (innerValue <= 0)
                            {
                                throw new ArgumentException("Cannot calculate logarithm of zero or negative number");
                            }
                            double logValue = Math.Log10(innerValue);
                            
                            string replacement = "(" + logValue.ToString($"F{decimalPlaces}", CultureInfo.InvariantCulture) + ")";
                            expression = expression.Substring(0, logIdx) + replacement + expression.Substring(closeIdx);
                            
                            return PrepareExpressionForEvaluation(expression); // Return modified string
                        }
                    }

                    // Process ln functions (natural logarithm)
                    int lnIdx = expression.IndexOf("ln(");
                    if (lnIdx >= 0)
                    {
                        int openCount = 1;
                        int closeIdx = lnIdx + 3;  // Start after "ln("
                        
                        while (openCount > 0 && closeIdx < expression.Length)
                        {
                            if (expression[closeIdx] == '(') openCount++;
                            if (expression[closeIdx] == ')') openCount--;
                            closeIdx++;
                        }
                        
                        if (openCount == 0)
                        {
                            string innerExpr = expression.Substring(lnIdx + 3, closeIdx - lnIdx - 4);
                            
                            if (innerExpr.Contains("sqrt(") || innerExpr.Contains("sin(") || 
                                innerExpr.Contains("cos(") || innerExpr.Contains("tan(") ||
                                innerExpr.Contains("log(") || innerExpr.Contains("ln("))
                            {
                                innerExpr = PrepareExpressionForEvaluation(innerExpr); // Recursive call
                            }
                            
                            DataTable dt = new DataTable();
                            double innerValue = Convert.ToDouble(dt.Compute(innerExpr, ""), CultureInfo.InvariantCulture);
                            if (innerValue <= 0)
                            {
                                throw new ArgumentException("Cannot calculate natural logarithm of zero or negative number");
                            }
                            double lnValue = Math.Log(innerValue);  // Natural logarithm
                            
                            string replacement = "(" + lnValue.ToString($"F{decimalPlaces}", CultureInfo.InvariantCulture) + ")";
                            expression = expression.Substring(0, lnIdx) + replacement + expression.Substring(closeIdx);
                            
                            return PrepareExpressionForEvaluation(expression); // Return modified string
                        }
                    }
                    
                    // If we get here without finding any functions to process, break to avoid infinite loop
                    break; 
                }
            }
            catch (ArgumentException ex) when (ex.Message == "NegativeSqrt")
            {
                throw; // Re-throw to be caught by CalculateResult
            }
            catch (Exception ex)
            {
                // Log the error for debugging or convert to a specific calculation error
                System.Diagnostics.Debug.WriteLine("Error in PrepareExpressionForEvaluation: " + ex.Message);
                throw new EvaluateException("Error evaluating sub-expression: " + expression, ex);
            }
            
            return expression; // Return the fully processed string
        }

        private bool HasBalancedParentheses(string expression)
        {
            int count = 0;
            foreach (char c in expression)
            {
                if (c == '(') count++;
                else if (c == ')') count--;

                // If at any point we have more closing than opening parentheses, it's invalid
                if (count < 0) return false;
            }
            // If count is 0, parentheses are balanced
            return count == 0;
        }

        private void ScientificButton_Click(object sender, EventArgs e)
        {
            Button button = (Button)sender;
            string operation = button.Tag.ToString();

            // Handle constants
            if (operation == "π")
            {
                if (isNewCalculation || string.IsNullOrEmpty(expressionDisplayTextBox.Text)) // was displayTextBox.Text == "0"
                {
                    expressionDisplayTextBox.Text = "π";  // Changed from PI.ToString() to "π"
                }
                else
                {
                    expressionDisplayTextBox.Text += "π";  // Changed from PI.ToString() to "π"
                }
                isNewCalculation = false;
                SetCursorToEnd();
                return;
            }
            else if (operation == "e")
            {
                if (isNewCalculation || string.IsNullOrEmpty(expressionDisplayTextBox.Text)) // was displayTextBox.Text == "0"
                {
                    expressionDisplayTextBox.Text = "e";  // Changed from E.ToString() to "e"
                }
                else
                {
                    expressionDisplayTextBox.Text += "e";  // Changed from E.ToString() to "e"
                }
                isNewCalculation = false;
                SetCursorToEnd();
                return;
            }
            else if (operation == "deg_to_rad") // RAD button clicked
            {
                // Toggle to Radians mode if we're not already in it
                isRadiansMode = true;
                UpdateAngleModeButtonsDisplay();

                // If there's a value in the result field, also convert it from DEG to RAD
                // This conversion is illustrative; direct mode change is primary.
                if (lastAnswer is double currentDoubleVal)
                {
                    lastAnswer = currentDoubleVal * DEG_TO_RAD;
                    resultDisplayTextBox.Text = ((double)lastAnswer).ToString($"F{decimalPlaces}", CultureInfo.InvariantCulture);
                    // expressionDisplayTextBox.Text = resultDisplayTextBox.Text; // Optional: update expression display
                }
                // else if lastAnswer is Complex, conversion is more involved. For now, only convert if double.
                return;
            }
            else if (operation == "rad_to_deg") // DEG button clicked
            {
                // Toggle to Degrees mode if we're not already in it
                isRadiansMode = false;
                UpdateAngleModeButtonsDisplay();

                if (lastAnswer is double currentDoubleVal)
                {
                    lastAnswer = currentDoubleVal * RAD_TO_DEG;
                    resultDisplayTextBox.Text = ((double)lastAnswer).ToString($"F{decimalPlaces}", CultureInfo.InvariantCulture);
                }
                return;
            }

            // Handle functions that require parentheses
            if (operation == "sqrt" || operation == "√" || operation == "sin" || operation == "cos" || 
                operation == "tan" || operation == "log" || operation == "ln")
            {
                // Start or continue an expression with this function
                if (isNewCalculation || string.IsNullOrEmpty(expressionDisplayTextBox.Text)) // was displayTextBox.Text == "0"
                {
                    if (operation == "sqrt" || operation == "√")
                        expressionDisplayTextBox.Text = "√(";
                    else
                        expressionDisplayTextBox.Text = operation + "(";
                }
                else
                {
                    if (operation == "sqrt" || operation == "√")
                        expressionDisplayTextBox.Text += "√(";
                    else
                        expressionDisplayTextBox.Text += operation + "(";
                }

                isExpressionMode = true;
                bracketCount++;
                isNewCalculation = false;
                SetCursorToEnd();
                return;
            }
            
            // Special handling for power function
            else if (operation == "x^y")
            {
                if (isNewCalculation || string.IsNullOrEmpty(expressionDisplayTextBox.Text)) // was displayTextBox.Text == "0"
                {
                    // Can't start with power operator if expression is empty
                    // Prepend with last answer or 0
                    expressionDisplayTextBox.Text = resultDisplayTextBox.Text + "^";
                } else {
                     expressionDisplayTextBox.Text += "^";
                }
                
                isExpressionMode = true; // Treat power operations as expressions
                isNewCalculation = false;
                SetCursorToEnd();
                return;
            }

            // For other operations like 1/x
            if (isExpressionMode)
            {
                // In expression mode, just append the function
                AppendToDisplay(operation);
            }
            else
            {
                // Try traditional calculation if possible, using the result display
                if (lastAnswer is double value) // Use lastAnswer for 1/x
                {
                    double result = 0;
                    bool calculated = false;

                    switch (operation)
                    {
                        case "1/x":
                            if (value != 0)
                            {
                                result = 1 / value;
                                calculated = true;
                            }
                            else
                            {
                                MessageBox.Show("Cannot divide by zero", "Error",
                                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                            }
                            break;
                    }

                    if (calculated)
                    {
                        lastAnswer = result;
                        resultDisplayTextBox.Text = result.ToString($"F{decimalPlaces}", CultureInfo.InvariantCulture); 
                        expressionDisplayTextBox.Text = resultDisplayTextBox.Text; 
                        isNewCalculation = true;
                    }
                    else
                    {
                        AppendToDisplay(operation);
                    }
                }
                else if (lastAnswer is Complex)
                {
                     // Handle 1/x for complex if needed, e.g., Complex.Reciprocal((Complex)lastAnswer)
                     // For now, just append to expression or show error
                    AppendToDisplay(operation); // Or show error: "1/x on complex not fully supported"
                }
                else
                {
                    AppendToDisplay(operation);
                }
            }
        }

        private void MemoryButton_Click(object sender, EventArgs e)
        {
            Button button = (Button)sender;
            string operation = button.Tag.ToString();
            double value;

            switch (operation)
            {
                case "MC":
                    memory = 0;
                    break;
                case "MR":
                    // When recalling memory, it should go into the expression display
                    // and also update the result display if it's a new calculation.
                    if (isNewCalculation) {
                        expressionDisplayTextBox.Text = memory.ToString();
                        resultDisplayTextBox.Text = memory.ToString();
                    } else {
                        expressionDisplayTextBox.Text += memory.ToString();
                    }
                    isNewCalculation = false; // Continue expression
                    break;
                case "M+":
                    if (double.TryParse(resultDisplayTextBox.Text, out value)) // Use result display for M+
                    {
                        memory += value;
                    }
                    isNewCalculation = true; // M+ usually finalizes current number
                    break;
                case "M-":
                    if (double.TryParse(resultDisplayTextBox.Text, out value)) // Use result display for M-
                    {
                        memory -= value;
                    }
                    isNewCalculation = true; // M- usually finalizes current number
                    break;
            }
        }

        private void ClearButton_Click(object sender, EventArgs e)
        {
            expressionDisplayTextBox.Text = ""; // Clear expression
            resultDisplayTextBox.Text = "0";    // Reset result to 0
            lastValue = 0;
            lastOperation = "";
            isNewCalculation = true;
            isExpressionMode = false;
            bracketCount = 0;  // Reset bracket count if you're tracking it
            keyBuffer = ""; // Reset key buffer
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
                        if (double.TryParse(resultDisplayTextBox.Text, out value)) // Insert from result display
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

        private void GetButton_Click(object sender, EventArgs e)
        {
            try
            {
                Excel.Application excelApp = Globals.ThisAddIn.Application;
                if (excelApp != null)
                {
                    Excel.Range activeCell = excelApp.ActiveCell;
                    if (activeCell != null)
                    {
                        // Get the value from the active cell
                        object cellValue = activeCell.Value;
                        if (cellValue != null)
                        {
                            double numericValue;
                            
                            // Special handling for date/time values
                            if (cellValue is DateTime dateTimeValue)
                            {
                                // Convert DateTime to Excel's numeric representation (days since 1/1/1900)
                                numericValue = dateTimeValue.ToOADate();
                            }
                            // Try to parse any string representation into a number
                            else if (!double.TryParse(cellValue.ToString(), out numericValue))
                            {
                                // For non-numeric values, try to parse as DateTime
                                DateTime parsedDateTime;
                                if (DateTime.TryParse(cellValue.ToString(), out parsedDateTime))
                                {
                                    numericValue = parsedDateTime.ToOADate();
                                }
                                else
                                {
                                    MessageBox.Show("The selected cell doesn't contain a value that can be converted to a number.", 
                                                  "Non-numeric Value", MessageBoxButtons.OK, MessageBoxIcon.Information);
                                    return;
                                }
                            }

                            // If we're adding to an existing expression
                            if (!isNewCalculation && !string.IsNullOrEmpty(expressionDisplayTextBox.Text)) 
                            {
                                // Append the value to the current expression
                                expressionDisplayTextBox.Text += numericValue.ToString(CultureInfo.InvariantCulture);
                            }
                            // If we're starting a new calculation
                            else
                            {
                                // Just set the value directly
                                expressionDisplayTextBox.Text = numericValue.ToString(CultureInfo.InvariantCulture);
                                resultDisplayTextBox.Text = numericValue.ToString(CultureInfo.InvariantCulture);
                                isNewCalculation = false; // Don't reset - continue using this value
                            }

                            // Always set cursor to the end
                            SetCursorToEnd();
                        }
                        Marshal.ReleaseComObject(activeCell);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error getting cell value: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void SetCursorToEnd()
        {
            expressionDisplayTextBox.Focus(); // Ensure focus is on the expression box
            expressionDisplayTextBox.SelectionStart = expressionDisplayTextBox.Text.Length;
            expressionDisplayTextBox.SelectionLength = 0;
        }

        private void TestNestedExpression(string expression)
        {
            // Clear the calculator
            ClearButton_Click(this, EventArgs.Empty);
            
            // Set the display text manually for testing
            expressionDisplayTextBox.Text = expression;
            isNewCalculation = false;
            isExpressionMode = true;
            
            // Count the brackets in the expression
            bracketCount = 0;
            foreach (char c in expression)
            {
                if (c == '(') bracketCount++;
                else if (c == ')') bracketCount--;
            }
            
            // Try to calculate
            CalculateResult();
        }

        private void BackspaceButton_Click(object sender, EventArgs e)
        {
            // Reuse the same logic that's used for keyboard backspace
            if (expressionDisplayTextBox.Text.Length > 0)
            {
                // Check if we're deleting a parenthesis
                char lastChar = expressionDisplayTextBox.Text[expressionDisplayTextBox.Text.Length - 1];
                if (lastChar == '(')
                    bracketCount--;
                else if (lastChar == ')')
                    bracketCount++;

                expressionDisplayTextBox.Text = expressionDisplayTextBox.Text.Substring(0, expressionDisplayTextBox.Text.Length - 1);
                if (expressionDisplayTextBox.Text.Length == 0)
                {
                    // expressionDisplayTextBox.Text = "0"; // Reset to 0 if empty - leave empty
                    resultDisplayTextBox.Text = "0"; // Reset result display
                    isNewCalculation = true;
                }
                
                // After backspace, recalculate if we should be in expression mode
                isExpressionMode = bracketCount > 0 ||
                                  expressionDisplayTextBox.Text.Contains("(") && !HasBalancedParentheses(expressionDisplayTextBox.Text);
                
                // Only format if not in expression mode
                if (!isExpressionMode)
                {
                    // FormatDisplayText(); // Consider when to call
                }
                
                SetCursorToEnd();
            }
        }

        private void ClearEntryButton_Click(object sender, EventArgs e)
        {
            expressionDisplayTextBox.Text = ""; // Clear expression
            resultDisplayTextBox.Text = "0";    // Reset result to 0
            lastValue = 0; // Reset lastValue when clearing entry
            lastOperation = "";
            isNewCalculation = true;
            isExpressionMode = false;
            bracketCount = 0;  // Reset bracket count
            keyBuffer = ""; // Reset key buffer
            SetCursorToEnd();
        }

        private void LastAnsButton_Click(object sender, EventArgs e)
        {
            // Insert the last answer into the current expression
            if (isNewCalculation || string.IsNullOrEmpty(expressionDisplayTextBox.Text)) // was displayTextBox.Text == "0"
            {
                // If starting a new calculation, replace with last answer
                expressionDisplayTextBox.Text = lastAnswer.ToString();
                isNewCalculation = false;
            }
            else
            {
                // Otherwise append to the current expression
                expressionDisplayTextBox.Text += lastAnswer.ToString();
            }
            
            // After appending, check if expression mode should be active
            isExpressionMode = bracketCount > 0 || 
                              (expressionDisplayTextBox.Text.Contains("(") && !HasBalancedParentheses(expressionDisplayTextBox.Text));
            SetCursorToEnd();
        }

        private void DebugExpression(string expression)
        {
            try
            {
                // Test the expression in isolation
                DataTable dt = new DataTable();
                var result = dt.Compute(expression, "");
                MessageBox.Show($"Expression: {expression}\nResult: {result}");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Expression error: {expression}\n{ex.Message}");
            }
        }

        // Add F1 key handler
        protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
        {
            if (keyData == Keys.F1)
            {
                ShowHelp();
                return true;
            }
            return base.ProcessCmdKey(ref msg, keyData);
        }

        private void HelpButton_Click(object sender, EventArgs e)
        {
            ShowHelp();
        }

        private void ShowHelp()
        {
            using (var helpForm = new CalculatorHelpForm())
            {
                helpForm.ShowDialog(this);
            }
        }

        // Update this method to save settings
        private void SettingsButton_Click(object sender, EventArgs e)
        {
            using (var settingsForm = new CalculatorSettingsForm())
            {
                settingsForm.Owner = this;  // Set the calculator as the owner
                settingsForm.DecimalPlaces = this.decimalPlaces;
                if (settingsForm.ShowDialog(this) == DialogResult.OK)  // Use ShowDialog(this)
                {
                    this.decimalPlaces = settingsForm.DecimalPlaces;
                    
                    // Save the setting when it changes
                    RowHighligher.Properties.Settings.Default.CalculatorDecimalPlaces = this.decimalPlaces;
                    RowHighligher.Properties.Settings.Default.Save();
                    
                    // If there's a number currently displayed in the result, reformat it
                    if (double.TryParse(resultDisplayTextBox.Text, out double currentValue)) // Check resultDisplayTextBox
                    {
                        resultDisplayTextBox.Text = currentValue.ToString($"F{decimalPlaces}"); // Update resultDisplayTextBox
                    }
                }
            }
        }

        // Method to update the display of RAD/DEG buttons based on current mode
        private void UpdateAngleModeButtonsDisplay()
        {
            Color activeColor = Color.LightSkyBlue;
            Color defaultColor = SystemColors.Control;

            // Find the RAD and DEG buttons using their tag values
            Button radButton = null;
            Button degButton = null;
            
            foreach (Button btn in scientificButtons)
            {
                if (btn.Tag.ToString() == "deg_to_rad")
                    radButton = btn;
                else if (btn.Tag.ToString() == "rad_to_deg")
                    degButton = btn;
            }

            if (radButton != null && degButton != null)
            {
                // Set colors based on current mode
                radButton.BackColor = isRadiansMode ? activeColor : defaultColor;
                degButton.BackColor = isRadiansMode ? defaultColor : activeColor;
            }
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

        protected override void OnActivated(EventArgs e)
        {
            base.OnActivated(e);
            // Update title bar colors
            customTitleBar.BackColor = Color.LightSkyBlue;
            closeButton.BackColor = Color.LightSkyBlue;
            minimizeButton.BackColor = Color.LightSkyBlue;
            
            // Set main panel background
            mainPanel.BackColor = Color.DarkGray;
            
            // Number buttons
            foreach (Button btn in numberButtons)
            {
                if (btn != null)
                    btn.BackColor = SystemColors.Control;
            }
            
            // Memory buttons - access them directly from the numberPanel
            for (int i = 0; i < 4; i++) // Memory buttons are in the first row (0) and columns 0-3
            {
                Control control = numberPanel.GetControlFromPosition(i, 0);
                if (control is Button btn)
                {
                    btn.BackColor = SystemColors.Control;
                }
            }
            
            // Parentheses panel buttons
            foreach (Control control in mainPanel.Controls)
            {
                if (control is TableLayoutPanel panel && panel != numberPanel)
                {
                    foreach (Control panelControl in panel.Controls)
                    {
                        if (panelControl is Button btn)
                        {
                            // Special handling for parentheses and other buttons
                            if (btn == helpButton || btn.Tag?.ToString() == "(" || 
                                btn.Tag?.ToString() == ")" || btn.Tag?.ToString() == "backspace" ||
                                btn.Tag?.ToString() == "clearEntry")
                            {
                                btn.BackColor = SystemColors.Control;
                            }
                            // Top row buttons (Get, Insert, LastAns, Settings)
                            else if (btn == insertButton || btn == clearButton || btn == settingsButton || 
                                     btn.Text == "Get")
                            {
                                btn.BackColor = SystemColors.Control;
                            }
                        }
                    }
                }
            }
            
            // Operator buttons with special colors
            for (int i = 0; i < operatorButtons.Length; i++)
            {
                if (operatorButtons[i] != null)
                {
                    if (i == 4) // Equals button
                        operatorButtons[i].BackColor = Color.Orange;
                    else if (i == 0 || i == 1 || i == 2 || i == 5) // +, -, *, / buttons
                        operatorButtons[i].BackColor = Color.LightGoldenrodYellow;
                    else // Other operator buttons (like decimal point)
                        operatorButtons[i].BackColor = SystemColors.Control;
                }
            }
            
            // Ensure scientific panel keeps its background
            foreach (Button btn in scientificButtons)
            {
                if (btn != null)
                    btn.BackColor = SystemColors.Control;
            }
            
            // Handle angle mode buttons
            UpdateAngleModeButtonsDisplay();
        }

        protected override void OnDeactivate(EventArgs e)
        {
            base.OnDeactivate(e);
            // Update title bar colors
            customTitleBar.BackColor = Color.FromArgb(210, 230, 250);
            closeButton.BackColor = Color.FromArgb(210, 230, 250);
            minimizeButton.BackColor = Color.FromArgb(210, 230, 250);
            
            // Reset main panel background
            mainPanel.BackColor = SystemColors.Control;
            
            // The buttons will keep their explicitly set colors from OnActivated
        }
    }

    // Custom exception for evaluation errors if needed
    public class EvaluateException : Exception
    {
        public EvaluateException(string message) : base(message) { }
        public EvaluateException(string message, Exception inner) : base(message, inner) { }
    }

    // Create a new class for the settings form
    public class CalculatorSettingsForm : Form
    {
        private NumericUpDown decimalPlacesInput;
        private int _decimalPlaces; // Renamed to avoid conflict with property

        public int DecimalPlaces
        {
            get { return _decimalPlaces; }
            set 
            { 
                _decimalPlaces = value; // Use backing field
                // Properties.Settings.Default.CalculatorDecimalPlaces = value; // Saving is done on Save button
                // Properties.Settings.Default.Save();
            }
        }

        public CalculatorSettingsForm()
        {
            this.Text = "General Settings";
            this.Size = new Size(400, 200);
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.StartPosition = FormStartPosition.CenterParent;
            this.TopMost = true;  // Add this line to make settings always on top

            TableLayoutPanel mainPanel = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                Padding = new Padding(10),
                RowCount = 2,
                ColumnCount = 2
            };

            Label decimalLabel = new Label
            {
                Text = "Choose decimal places:",
                AutoSize = true,
                Anchor = AnchorStyles.Left | AnchorStyles.Right,
                TextAlign = ContentAlignment.MiddleLeft
            };

            // Load from settings when creating input
            decimalPlacesInput = new NumericUpDown
            {
                Minimum = 0,
                Maximum = 15,
                // Value = Properties.Settings.Default.CalculatorDecimalPlaces, // Set by DecimalPlaces property
                Width = 60
            };
            // _decimalPlaces = Properties.Settings.Default.CalculatorDecimalPlaces; // Set by DecimalPlaces property
            decimalPlacesInput.ValueChanged += (s, e) => _decimalPlaces = (int)decimalPlacesInput.Value;


            mainPanel.Controls.Add(decimalLabel, 0, 0);
            mainPanel.Controls.Add(decimalPlacesInput, 1, 0);

            Button saveButton = new Button
            {
                Text = "Save",
                DialogResult = DialogResult.OK, // This will close the form with OK
                Anchor = AnchorStyles.Right
            };
            
            // Save settings when clicking save (DialogResult.OK handles closing)
            // The actual saving of Properties.Settings.Default.CalculatorDecimalPlaces
            // should happen in ScientificCalculator form after ShowDialog returns OK.
            // saveButton.Click += (s, e) => 
            // { 
            //     Properties.Settings.Default.CalculatorDecimalPlaces = _decimalPlaces;
            //     Properties.Settings.Default.Save();
            // };
            
            mainPanel.Controls.Add(saveButton, 1, 1);
            this.Controls.Add(mainPanel);
            
            // Ensure the NumericUpDown reflects the initial DecimalPlaces value
            // This should be done after the DecimalPlaces property is set by the caller.
            this.Load += (s, e) => decimalPlacesInput.Value = Math.Max(0, Math.Min(10, _decimalPlaces));
        }
    }
    
    public static class ExpressionEvaluator
    {
        // Main evaluation entry point
        public static object Evaluate(string expression, int decimalPlaces, bool isRadians)
        {
            try
            {
                // Initial replacements for constants
                expression = expression.Replace("π", Math.PI.ToString(CultureInfo.InvariantCulture));
                expression = expression.Replace("e", Math.E.ToString(CultureInfo.InvariantCulture));
                expression = expression.Replace(" ", "").Replace("×", "*").Replace("÷", "/");


                // Handle powers (^) - this simple replacement might need a proper parser for complex bases
                while (expression.Contains("^"))
                {
                    // This power handling is basic and assumes real numbers from DataTable.Compute
                    // It won't correctly handle complex bases or exponents without a full parser.
                    // For now, it might lead to errors if complex numbers are involved before power.
                    // For now, it might lead to errors if complex numbers are involved before power.
                    int powerIndex = expression.IndexOf("^");
                    int baseStartIndex = powerIndex - 1;
                    while (baseStartIndex >= 0 && (char.IsDigit(expression[baseStartIndex]) || expression[baseStartIndex] == '.' || expression[baseStartIndex] == ')' || expression[baseStartIndex] == '-')) // Added '-' for negative base
                        baseStartIndex--;
                    baseStartIndex++;
                    
                    int exponentEndIndex = powerIndex + 1;
                    // Allow negative exponents and parentheses in exponent
                    if (exponentEndIndex < expression.Length && expression[exponentEndIndex] == '-') exponentEndIndex++; 
                    while (exponentEndIndex < expression.Length && (char.IsDigit(expression[exponentEndIndex]) || expression[exponentEndIndex] == '.' || expression[exponentEndIndex] == '(' || expression[exponentEndIndex] == ')'))
                        exponentEndIndex++;
                    
                    string baseStr = expression.Substring(baseStartIndex, powerIndex - baseStartIndex);
                    string exponentStr = expression.Substring(powerIndex + 1, exponentEndIndex - (powerIndex + 1));

                    // Evaluate base and exponent - these might need to handle complex numbers if we go further
                    double baseValue = Convert.ToDouble(new DataTable().Compute(baseStr, ""), CultureInfo.InvariantCulture);
                    double exponentValue = Convert.ToDouble(new DataTable().Compute(exponentStr, ""), CultureInfo.InvariantCulture);
                    
                    double powerResult = Math.Pow(baseValue, exponentValue);
                    string replacement = "(" + powerResult.ToString("F" + decimalPlaces, CultureInfo.InvariantCulture) + ")"; // Ensure enough precision
                    expression = expression.Substring(0, baseStartIndex) + replacement + expression.Substring(exponentEndIndex);
                }

                // Process functions and final computation
                return EvaluateFunctionsRecursive(expression, decimalPlaces, isRadians);
            }
            catch (ArgumentException ex) when (ex.Message == "NegativeSqrt")
            {
                throw; // Re-throw for ScientificCalculator to catch and format Complex
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Evaluation Error: {ex}");
                throw new EvaluateException("Error evaluating expression: " + ex.Message, ex);
            }
        }

        // Recursive function evaluator
        private static object EvaluateFunctionsRecursive(string expression, int decimalPlaces, bool isRadians)
        {
            string[] functions = { "sqrt", "sin", "cos", "tan", "log", "ln" }; // Order might matter for parsing
            
            foreach (var funcName in functions)
            {
                int idx = -1;
                while ((idx = expression.IndexOf(funcName + "(", StringComparison.OrdinalIgnoreCase)) != -1)
                {
                    int openParenIndex = idx + funcName.Length;
                    int balance = 1;
                    int closeParenIndex = -1;

                    for (int i = openParenIndex + 1; i < expression.Length; i++)
                    {
                        if (expression[i] == '(') balance++;
                        else if (expression[i] == ')') balance--;
                        if (balance == 0)
                        {
                            closeParenIndex = i;
                            break;
                        }
                    }

                    if (closeParenIndex == -1)
                    {
                        throw new EvaluateException($"Mismatched parentheses for function {funcName}");
                    }

                    string innerExpr = expression.Substring(openParenIndex + 1, closeParenIndex - (openParenIndex + 1));
                    object innerResultObj = EvaluateFunctionsRecursive(innerExpr, decimalPlaces, isRadians); // Recursive call

                    double innerDoubleVal;
                    if (innerResultObj is Complex)
                    {
                        // Current functions (sin, cos, etc.) are not set up for Complex input here.
                        // This path would require Complex versions of all functions.
                        // For this limited sqrt implementation, we assume innerResult for sqrt will be double.
                        // If sqrt itself gets a complex number, Complex.Sqrt should be used.
                        if (funcName == "sqrt") {
                            return Complex.Sqrt((Complex)innerResultObj);
                        }
                        throw new EvaluateException($"Function '{funcName}' does not support complex arguments in this version.");
                    }
                    else if (innerResultObj is double val)
                    {
                        innerDoubleVal = val;
                    }
                    else
                    {
                        throw new EvaluateException($"Unexpected inner result type for function {funcName}: {innerResultObj?.GetType()}");
                    }

                    object funcResult = null;
                    switch (funcName.ToLowerInvariant())
                    {
                        case "sqrt":
                            if (innerDoubleVal < 0)
                            {
                                Complex complexSqrtResult = Complex.Sqrt(new Complex(innerDoubleVal, 0));
                                // Instead of direct return, throw specific exception for CalculateResult to handle
                                ArgumentException ex = new ArgumentException("NegativeSqrt");
                                ex.Data["ComplexResult"] = complexSqrtResult;
                                throw ex;
                            }
                            funcResult = Math.Sqrt(innerDoubleVal);
                            break;
                        case "sin":
                            funcResult = Math.Sin(isRadians ? innerDoubleVal : innerDoubleVal * CalculatorConstants.DEG_TO_RAD);
                            break;
                        case "cos":
                            funcResult = Math.Cos(isRadians ? innerDoubleVal : innerDoubleVal * CalculatorConstants.DEG_TO_RAD);
                            break;
                        case "tan":
                            funcResult = Math.Tan(isRadians ? innerDoubleVal : innerDoubleVal * CalculatorConstants.DEG_TO_RAD);
                            break;
                        case "log":
                            if (innerDoubleVal <= 0) throw new EvaluateException("Logarithm argument must be positive.");
                            funcResult = Math.Log10(innerDoubleVal);
                            break;
                        case "ln":
                            if (innerDoubleVal <= 0) throw new EvaluateException("Natural logarithm argument must be positive.");
                            funcResult = Math.Log(innerDoubleVal);
                            break;
                    }
                    
                    // Replace the function call with its result (as a string for further parsing)
                    string replacement;
                    if (funcResult is double dRes) {
                        replacement = "(" + dRes.ToString("G17", CultureInfo.InvariantCulture) + ")"; // Use G17 for precision
                    } else if (funcResult is Complex cRes) { // Should not happen with current limited scope
                        replacement = "(" + cRes.Real.ToString("G17", CultureInfo.InvariantCulture) + "+" + cRes.Imaginary.ToString("G17", CultureInfo.InvariantCulture) + "i)";
                    } else {
                         throw new EvaluateException("Unknown function result type.");
                    }
                    
                    expression = expression.Substring(0, idx) + replacement + expression.Substring(closeParenIndex + 1);
                    // After a replacement, re-evaluate the modified expression from the start of this level
                    // to handle nested functions or functions appearing multiple times correctly.
                    // This continues the loop for the current funcName on the modified expression.
                }
            }
            
            // After all functions are processed, the expression should be evaluatable by DataTable.Compute
            // if it resolved to a simple arithmetic string of real numbers.
            try
            {
                return Convert.ToDouble(new DataTable().Compute(expression, ""), CultureInfo.InvariantCulture);
            }
            catch (SyntaxErrorException e) // Catch specific DataTable.Compute errors
            {
                // This can happen if the expression is malformed after function replacements
                // or if it contained an unhandled complex number representation.
                throw new EvaluateException("Invalid expression for final computation: " + expression, e);
            }
        }
    }
}
