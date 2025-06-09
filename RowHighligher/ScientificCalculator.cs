using System;
using System.Drawing;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;
using System.Data;
using System.Text;

namespace RowHighligher
{
    public partial class ScientificCalculator : Form
    {
        private TextBox displayTextBox;
        private Button insertButton;
        private Button clearButton;
        private Button helpButton;  // Add help button field
        private Button settingsButton; // Add settings button field
        private Button[] numberButtons;
        private Button[] operatorButtons;
        private Button[] scientificButtons;
        private double memory = 0;
        private bool isNewCalculation = true;
        private string lastOperation = "";
        private double lastValue = 0;
        private bool isInKeyboardMode = true;

        // Add a field to track when an expression is being built
        private bool isExpressionMode = false;

        // Add this field near the top of your class with other fields
        private double lastAnswer = 0;

        // Add this field to store characters as they're typed
        private string keyBuffer = "";

        // Add constants
        private const double PI = Math.PI;
        private const double E = Math.E;
        private const double RAD_TO_DEG = 180.0 / Math.PI;
        private const double DEG_TO_RAD = Math.PI / 180.0;

        // Update the field to use the saved setting
        private int decimalPlaces = Properties.Settings.Default.CalculatorDecimalPlaces;

        public ScientificCalculator()
        {
            InitializeComponents();

            // Set form properties
            this.FormBorderStyle = FormBorderStyle.SizableToolWindow;
            this.Text = "Scientific Calculator";
            this.MinimumSize = new Size(350, 550);  // Increased minimum size
            this.Size = new Size(350, 600);         // Set larger default size
            this.MaximizeBox = false;
            this.MinimizeBox = true;
            this.TopMost = Properties.Settings.Default.IsCalculatorDetached;
            this.StartPosition = FormStartPosition.CenterScreen;
            this.KeyPreview = true;

            // Add keyboard event handlers
            this.KeyDown += ScientificCalculator_KeyDown;
            this.KeyPress += ScientificCalculator_KeyPress;
            this.displayTextBox.GotFocus += DisplayTextBox_GotFocus;
            this.displayTextBox.LostFocus += DisplayTextBox_LostFocus;
        }

        // Add this to enhance the display's readability when showing expressions
        private void FormatDisplayText()
        {
            string text = displayTextBox.Text;
            
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
            
            displayTextBox.Text = text.Trim();
            
            // Always move cursor to end after formatting
            SetCursorToEnd();
        }

        private void InitializeComponents()
        {
            // Main layout panel
            TableLayoutPanel mainPanel = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                RowCount = 6,         // Reduced to 6 rows
                ColumnCount = 1,
                Padding = new Padding(10)
            };

            // Update row styles for better proportions
            mainPanel.RowStyles.Add(new RowStyle(SizeType.Percent, 10));  // Display
            mainPanel.RowStyles.Add(new RowStyle(SizeType.Percent, 10));  // Buttons panel
            mainPanel.RowStyles.Add(new RowStyle(SizeType.Percent, 10));  // Parentheses panel
            mainPanel.RowStyles.Add(new RowStyle(SizeType.Percent, 21));  // Scientific panel (increased)
            mainPanel.RowStyles.Add(new RowStyle(SizeType.Percent, 30));  // Number panel
            mainPanel.RowStyles.Add(new RowStyle(SizeType.Percent, 10));  // Bottom panel

            // Display textbox
            displayTextBox = new TextBox
            {
                Dock = DockStyle.Fill,
                ReadOnly = true,
                Font = new Font("Consolas", 20, FontStyle.Bold),
                TextAlign = HorizontalAlignment.Right,
                Text = "0"
            };

            // Add panel for display and decimal places control
            Panel displayPanel = new Panel
            {
                Dock = DockStyle.Fill,
                Padding = new Padding(0)
            };
            displayPanel.Controls.Add(displayTextBox);

            mainPanel.Controls.Add(displayPanel, 0, 0);

            // Insert button in a panel with Clear and decimal places
            TableLayoutPanel buttonsPanel = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 5,  // Increase column count to add logo
                RowCount = 1
            };
            for (int i = 0; i < 6; i++) // Update column count
            {
                buttonsPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, i == 4 ? 28 : 18)); // Logo column bigger
            }

            // Add help button
            helpButton = new Button
            {
                Text = "Help (F1)",
                Dock = DockStyle.Fill,
                Font = new Font("Segoe UI", 8.25f, FontStyle.Regular),
                Margin = new Padding(1)
            };
            helpButton.Click += HelpButton_Click;
            buttonsPanel.Controls.Add(helpButton, 0, 0);

            insertButton = new Button
            {
                Text = "Insert (Ctrl+Enter)",
                Dock = DockStyle.Fill,
                Font = new Font("Segoe UI", 8.25f, FontStyle.Regular),
                Margin = new Padding(1)
            };
            insertButton.Click += InsertButton_Click;
            buttonsPanel.Controls.Add(insertButton, 1, 0);

            clearButton = new Button
            {
                Text = "LastAns",
                Dock = DockStyle.Fill,
                Font = new Font("Segoe UI", 8.25f, FontStyle.Regular),
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
                Image = Properties.Resources.Bankers_Logo_Albania // Add your PNG image to resources
            };

            // Add logo after settings button
            buttonsPanel.Controls.Add(logoBox, 4, 0);

            // Add settings button
            settingsButton = new Button
            {
                Text = "Settings",
                Dock = DockStyle.Fill,
                Font = new Font("Segoe UI", 8.25f, FontStyle.Regular),
                Margin = new Padding(1)
            };
            settingsButton.Click += SettingsButton_Click;
            buttonsPanel.Controls.Add(settingsButton, 3, 0);

            mainPanel.Controls.Add(buttonsPanel, 0, 1);

            // Parentheses panel
            TableLayoutPanel parenthesesPanel = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 4, // Changed from 2 to 4
                RowCount = 1
            };
            parenthesesPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 25));
            parenthesesPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 25));
            parenthesesPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 25));
            parenthesesPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 25));

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

            // Add the parentheses panel to the main panel
            mainPanel.Controls.Add(parenthesesPanel, 0, 2);

            // Adjust the row index for subsequent panels
            //mainPanel.Controls.Add(scientificPanel, 0, 3);

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
            mainPanel.Controls.Add(scientificBorderPanel, 0, 3);  // Add to row 3

            // Number buttons panel for rows 3, 4, 5
            TableLayoutPanel numberPanel = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 4,
                RowCount = 4,
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

            // Number buttons 7-9 and /
            numberButtons = new Button[10]; // Digits 0-9
            operatorButtons = new Button[8]; // +, -, *, /, =, ., +/-, ?

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

            operatorButtons[0] = new Button { Text = "/", Dock = DockStyle.Fill, Tag = "/" };
            operatorButtons[0].Click += OperatorButton_Click;
            numberPanel.Controls.Add(operatorButtons[0], 3, 1);

            // Number buttons 4-6 and *
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

            operatorButtons[1] = new Button { Text = "*", Dock = DockStyle.Fill, Tag = "*" };
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

            operatorButtons[2] = new Button { Text = "-", Dock = DockStyle.Fill, Tag = "-" };
            operatorButtons[2].Click += OperatorButton_Click;
            numberPanel.Controls.Add(operatorButtons[2], 3, 3);

            mainPanel.Controls.Add(numberPanel, 0, 4);  // Add to row 4 (moved down)

            // Bottom panel for 0, ., +/-, =, +
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

            operatorButtons[4] = new Button { Text = "=", Dock = DockStyle.Fill, Tag = "=" };
            operatorButtons[4].Click += EqualsButton_Click;
            bottomPanel.Controls.Add(operatorButtons[4], 2, 0);

            operatorButtons[5] = new Button { Text = "+", Dock = DockStyle.Fill, Tag = "+" };
            operatorButtons[5].Click += OperatorButton_Click;
            bottomPanel.Controls.Add(operatorButtons[5], 3, 0);

            mainPanel.Controls.Add(bottomPanel, 0, 5);

            this.Controls.Add(mainPanel);
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
                if (isNewCalculation || displayTextBox.Text == "0")
                {
                    displayTextBox.Text = keyBuffer;
                }
                else
                {
                    displayTextBox.Text += e.KeyChar; // Add just the current character
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
                        string currentTextInner = displayTextBox.Text; // Renamed to 'currentTextInner'
                        displayTextBox.Text = currentTextInner.Substring(0, currentTextInner.Length - 2) + "π";
                        keyBuffer = "";
                        isNewCalculation = false;
                        SetCursorToEnd();
                        e.Handled = true;
                        return;
                    }
                    else if (keyBuffer.Equals("e", StringComparison.OrdinalIgnoreCase))
                    {
                        // Special handling for e
                        string currentTextInner = displayTextBox.Text; // Renamed to 'currentTextInner'
                        displayTextBox.Text = currentTextInner.Substring(0, currentTextInner.Length - 1) + "e";
                        keyBuffer = "";
                        isNewCalculation = false;
                        SetCursorToEnd();
                        e.Handled = true;
                        return;
                    }
                    
                    // Remove the function characters and add the function with parenthesis
                    string currentTextOuter = displayTextBox.Text;
                    displayTextBox.Text = currentTextOuter.Substring(0, currentTextOuter.Length - charsToCut) + 
                                          functionName + "(";
                    
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
                if (displayTextBox.Text.Length > 0)
                {
                    // Check if we're deleting a parenthesis
                    char lastChar = displayTextBox.Text[displayTextBox.Text.Length - 1];
                    if (lastChar == '(')
                        bracketCount--;
                    else if (lastChar == ')')
                        bracketCount++;

                    displayTextBox.Text = displayTextBox.Text.Substring(0, displayTextBox.Text.Length - 1);
                    if (displayTextBox.Text.Length == 0)
                    {
                        displayTextBox.Text = "0"; // Reset to 0 if empty
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
                displayTextBox.Text += "^";
                isExpressionMode = true; // Treat power operations as expressions
                isNewCalculation = false;
                SetCursorToEnd();
                e.Handled = true;
            }
        }

        private void DisplayTextBox_GotFocus(object sender, EventArgs e)
        {
            isInKeyboardMode = true;
            displayTextBox.BackColor = Color.LightSkyBlue;
        }

        private void DisplayTextBox_LostFocus(object sender, EventArgs e)
        {
            isInKeyboardMode = false;
            displayTextBox.BackColor = SystemColors.Window;
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
                // For starting a new calculation
                if (value == "sqrt" || value == "sin" || value == "cos")
                {
                    // For functions, show the function name followed by opening parenthesis
                    displayTextBox.Text = value + "(";
                    isExpressionMode = true;
                    bracketCount++; // Important: increment bracket count when adding function parenthesis
                }
                else if (value == "(")
                {
                    // For opening parenthesis
                    displayTextBox.Text = value;
                    bracketCount++;
                    isExpressionMode = true; // Important: set expression mode when opening parenthesis
                }
                else if (value == "-")
                {
                    // For negative number
                    displayTextBox.Text = value;
                }
                else if (char.IsDigit(value[0]) || value == ".")
                {
                    // For digits or decimal point
                    displayTextBox.Text = value;
                }
                else
                {
                    // For operators, start with 0 then operator
                    displayTextBox.Text = "0" + value;
                }
                isNewCalculation = false;
            }
            else
            {
                // For continuing an expression
                if (value == "sqrt" || value == "sin" || value == "cos")
                {
                    displayTextBox.Text += value + "(";
                    isExpressionMode = true;
                    bracketCount++; // Important: increment bracket count when adding function parenthesis
                }
                else if (value == "(")
                {
                    displayTextBox.Text += value;
                    bracketCount++;
                    isExpressionMode = true; // Important: set expression mode when opening parenthesis
                }
                else if (displayTextBox.Text == "0" && value != ".")
                {
                    displayTextBox.Text = value;
                }
                else
                {
                    displayTextBox.Text += value;
                }
            }

            // After appending the value, always check if we should be in expression mode
            // If there are any open brackets, we should be in expression mode
            isExpressionMode = bracketCount > 0 ||
                              displayTextBox.Text.Contains("(") && !HasBalancedParentheses(displayTextBox.Text);


            // Only format the display once, and only when appropriate
            if (!isExpressionMode)
            {
                FormatDisplayText();
            }

            // Always set cursor to the end
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
                displayTextBox.Text = "0.";
                isNewCalculation = false;
            }
            else
            {
                // Check if we need to handle multiple terms
                string currentText = displayTextBox.Text;
                
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
                            displayTextBox.Text += "0.";
                        else
                            displayTextBox.Text += ".";
                    }
                }
                else
                {
                    // Working with just the first term
                    if (!currentText.Contains("."))
                        displayTextBox.Text += ".";
                }
            }
            
            // Format and set cursor position
            FormatDisplayText(); 
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
            if (operation == "-" && (isNewCalculation || displayTextBox.Text == "0"))
            {
                // Special case for negative numbers at start
                displayTextBox.Text = operation;
                isNewCalculation = false;
            }
            else if (isNewCalculation)
            {
                // If starting fresh with non-negative operator, prepend 0
                displayTextBox.Text = "0 " + operation + " ";
                lastOperation = operation;
                isNewCalculation = false;
            }
            else
            {
                // For continuing an expression or after result
                displayTextBox.Text += " " + operation + " ";
                lastOperation = operation;
                isNewCalculation = false;
            }
            
            SetCursorToEnd();
        }

        private void EqualsButton_Click(object sender, EventArgs e)
        {
            CalculateResult();
            lastValue = lastAnswer; // Store the last calculated result in lastValue
        }

        private void CalculateResult()
        {
            try
            {
                if (isNewCalculation || displayTextBox.Text == "0")
                {
                    return; // No calculation needed
                }

                if (!HasBalancedParentheses(displayTextBox.Text))
                {
                    if (displayTextBox.Text.Contains("(") || displayTextBox.Text.Contains(")"))
                    {
                        MessageBox.Show("Expression has unbalanced parentheses. Please close all open brackets.",
                            "Syntax Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        bracketCount = CountOpenParentheses(displayTextBox.Text);
                        return;
                    }
                }

                isExpressionMode = false;
                bracketCount = 0;

                double result = ExpressionEvaluator.Evaluate(displayTextBox.Text, decimalPlaces);
                lastAnswer = result;
                displayTextBox.Text = result.ToString($"F{decimalPlaces}");

                lastOperation = "";
                isNewCalculation = true;
                SetCursorToEnd();
            }
            catch (Exception ex)
            {
                ShowError("Error calculating result: " + ex.Message);
            }
        }

        private void ShowError(string message)
        {
            displayTextBox.Text = "Error";
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
            expression = expression.Replace("π", PI.ToString());
            expression = expression.Replace("e", E.ToString());

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
                double baseValue = Convert.ToDouble(dt.Compute(baseStr, ""));
                double exponentValue = Convert.ToDouble(dt.Compute(exponentStr, ""));
                double powerResult = Math.Pow(baseValue, exponentValue);
                
                // Replace the power expression with the result
                string replacement = "(" + powerResult.ToString($"F{decimalPlaces}") + ")";
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
                                innerExpr = PrepareExpressionForEvaluation(innerExpr);
                            }
                            
                            // Calculate the sqrt of the inner expression
                            DataTable dt = new DataTable();
                            double innerValue = Convert.ToDouble(dt.Compute(innerExpr, ""));
                            double sqrtValue = Math.Sqrt(innerValue);
                            
                            // KEY FIX: Add parentheses around the result to ensure proper operator precedence
                            string replacement = "(" + sqrtValue.ToString($"F{decimalPlaces}") + ")";
                            expression = expression.Substring(0, sqrtIdx) + replacement + expression.Substring(closeIdx);
                            
                            // Since we've modified the expression, we need to process it again
                            return PrepareExpressionForEvaluation(expression);
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
                                innerExpr = PrepareExpressionForEvaluation(innerExpr);
                            }
                            
                            // Calculate sin of the inner expression
                            DataTable dt = new DataTable();
                            double innerValue = Convert.ToDouble(dt.Compute(innerExpr, ""));
                            double sinValue = Math.Sin(innerValue);
                            
                            // Replace with result (in parentheses)
                            string replacement = "(" + sinValue.ToString($"F{decimalPlaces}") + ")";
                            expression = expression.Substring(0, sinIdx) + replacement + expression.Substring(closeIdx);
                            
                            // Process the updated expression
                            return PrepareExpressionForEvaluation(expression);
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
                                innerExpr = PrepareExpressionForEvaluation(innerExpr);
                            }
                            
                            DataTable dt = new DataTable();
                            double innerValue = Convert.ToDouble(dt.Compute(innerExpr, ""));
                            double cosValue = Math.Cos(innerValue);
                            
                            string replacement = "(" + cosValue.ToString($"F{decimalPlaces}") + ")";
                            expression = expression.Substring(0, cosIdx) + replacement + expression.Substring(closeIdx);
                            
                            return PrepareExpressionForEvaluation(expression);
                        }
                    }

                    // Process tan functions
                    int tanIdx = expression.IndexOf("tan(");
                    if (tanIdx >= 0)
                    {
                        int openCount = 1;
                        int closeIdx = tanIdx + 4;  // Start after "tan(")
                        
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
                                innerExpr = PrepareExpressionForEvaluation(innerExpr);
                            }
                            
                            DataTable dt = new DataTable();
                            double innerValue = Convert.ToDouble(dt.Compute(innerExpr, ""));
                            double tanValue = Math.Tan(innerValue);
                            
                            string replacement = "(" + tanValue.ToString($"F{decimalPlaces}") + ")";
                            expression = expression.Substring(0, tanIdx) + replacement + expression.Substring(closeIdx);
                            
                            return PrepareExpressionForEvaluation(expression);
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
                                innerExpr = PrepareExpressionForEvaluation(innerExpr);
                            }
                            
                            DataTable dt = new DataTable();
                            double innerValue = Convert.ToDouble(dt.Compute(innerExpr, ""));
                            if (innerValue <= 0)
                            {
                                throw new ArgumentException("Cannot calculate logarithm of zero or negative number");
                            }
                            double logValue = Math.Log10(innerValue);
                            
                            string replacement = "(" + logValue.ToString($"F{decimalPlaces}") + ")";
                            expression = expression.Substring(0, logIdx) + replacement + expression.Substring(closeIdx);
                            
                            return PrepareExpressionForEvaluation(expression);
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
                                innerExpr = PrepareExpressionForEvaluation(innerExpr);
                            }
                            
                            DataTable dt = new DataTable();
                            double innerValue = Convert.ToDouble(dt.Compute(innerExpr, ""));
                            if (innerValue <= 0)
                            {
                                throw new ArgumentException("Cannot calculate natural logarithm of zero or negative number");
                            }
                            double lnValue = Math.Log(innerValue);  // Natural logarithm
                            
                            string replacement = "(" + lnValue.ToString($"F{decimalPlaces}") + ")";
                            expression = expression.Substring(0, lnIdx) + replacement + expression.Substring(closeIdx);
                            
                            return PrepareExpressionForEvaluation(expression);
                        }
                    }
                    
                    // If we get here without finding any functions to process, break to avoid infinite loop
                    break;
                }
            }
            catch (Exception ex)
            {
                // Log the error for debugging
                System.Diagnostics.Debug.WriteLine("Error evaluating expression: " + ex.Message);
            }
            
            return expression;
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
                if (isNewCalculation || displayTextBox.Text == "0")
                {
                    displayTextBox.Text = "π";  // Changed from PI.ToString() to "π"
                }
                else
                {
                    displayTextBox.Text += "π";  // Changed from PI.ToString() to "π"
                }
                isNewCalculation = false;
                SetCursorToEnd();
                return;
            }
            else if (operation == "e")
            {
                if (isNewCalculation || displayTextBox.Text == "0")
                {
                    displayTextBox.Text = "e";  // Changed from E.ToString() to "e"
                }
                else
                {
                    displayTextBox.Text += "e";  // Changed from E.ToString() to "e"
                }
                isNewCalculation = false;
                SetCursorToEnd();
                return;
            }
            else if (operation == "deg_to_rad" || operation == "rad_to_deg")
            {
                double value;
                if (double.TryParse(displayTextBox.Text, out value))
                {
                    if (operation == "deg_to_rad")
                    {
                        value *= DEG_TO_RAD;
                    }
                    else
                    {
                        value *= RAD_TO_DEG;
                    }
                    displayTextBox.Text = value.ToString();
                    isNewCalculation = true;
                }
                return;
            }

            // Handle functions that require parentheses
            if (operation == "sqrt" || operation == "sin" || operation == "cos" || 
                operation == "tan" || operation == "log" || operation == "ln")
            {
                // Start or continue an expression with this function
                if (isNewCalculation || displayTextBox.Text == "0")
                {
                    displayTextBox.Text = operation + "(";
                }
                else
                {
                    displayTextBox.Text += operation + "(";
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
                if (isNewCalculation || displayTextBox.Text == "0")
                {
                    // Can't start with power operator
                    return;
                }
                
                // Add the power operator
                displayTextBox.Text += "^";
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
                // Try traditional calculation if possible
                double value;
                if (double.TryParse(displayTextBox.Text, out value))
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
                        displayTextBox.Text = result.ToString();
                        isNewCalculation = true;
                    }
                    else
                    {
                        // If calculation failed, switch to expression mode
                        AppendToDisplay(operation);
                    }
                }
                else
                {
                    // If not a valid number, switch to expression mode
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
                    displayTextBox.Text = memory.ToString();
                    isNewCalculation = true;
                    break;
                case "M+":
                    if (double.TryParse(displayTextBox.Text, out value))
                    {
                        memory += value;
                    }
                    isNewCalculation = true;
                    break;
                case "M-":
                    if (double.TryParse(displayTextBox.Text, out value))
                    {
                        memory -= value;
                    }
                    isNewCalculation = true;
                    break;
            }
        }

        private void ClearButton_Click(object sender, EventArgs e)
        {
            displayTextBox.Text = "0";
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
                        if (double.TryParse(displayTextBox.Text, out value))
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

        private void SetCursorToEnd()
        {
            displayTextBox.SelectionStart = displayTextBox.Text.Length;
            displayTextBox.SelectionLength = 0;
        }

        private void TestNestedExpression(string expression)
        {
            // Clear the calculator
            ClearButton_Click(this, EventArgs.Empty);
            
            // Set the display text manually for testing
            displayTextBox.Text = expression;
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
            if (displayTextBox.Text.Length > 0)
            {
                // Check if we're deleting a parenthesis
                char lastChar = displayTextBox.Text[displayTextBox.Text.Length - 1];
                if (lastChar == '(')
                    bracketCount--;
                else if (lastChar == ')')
                    bracketCount++;

                displayTextBox.Text = displayTextBox.Text.Substring(0, displayTextBox.Text.Length - 1);
                if (displayTextBox.Text.Length == 0)
                {
                    displayTextBox.Text = "0"; // Reset to 0 if empty
                    isNewCalculation = true;
                }
                
                // After backspace, recalculate if we should be in expression mode
                isExpressionMode = bracketCount > 0 ||
                                  displayTextBox.Text.Contains("(") && !HasBalancedParentheses(displayTextBox.Text);
                
                // Only format if not in expression mode
                if (!isExpressionMode)
                {
                    FormatDisplayText();
                }
                
                SetCursorToEnd();
            }
        }

        private void ClearEntryButton_Click(object sender, EventArgs e)
        {
            displayTextBox.Text = "0";
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
            if (isNewCalculation || displayTextBox.Text == "0")
            {
                // If starting a new calculation, replace with last answer
                displayTextBox.Text = lastAnswer.ToString();
                isNewCalculation = false;
            }
            else
            {
                // Otherwise append to the current expression
                displayTextBox.Text += lastAnswer.ToString();
            }
            
            // After appending, check if expression mode should be active
            isExpressionMode = bracketCount > 0 || 
                              (displayTextBox.Text.Contains("(") && !HasBalancedParentheses(displayTextBox.Text));
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
                    Properties.Settings.Default.CalculatorDecimalPlaces = this.decimalPlaces;
                    Properties.Settings.Default.Save();
                    
                    // If there's a number currently displayed, reformat it with the new decimal places
                    if (double.TryParse(displayTextBox.Text, out double currentValue))
                    {
                        displayTextBox.Text = currentValue.ToString($"F{decimalPlaces}");
                    }
                }
            }
        }
    }

    // Create a new class for the settings form
    public class CalculatorSettingsForm : Form
    {
        private NumericUpDown decimalPlacesInput;
        private int decimalPlaces;

        public int DecimalPlaces
        {
            get { return decimalPlaces; }
            set 
            { 
                decimalPlaces = value;
                Properties.Settings.Default.CalculatorDecimalPlaces = value;
                Properties.Settings.Default.Save();
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
                Maximum = 10,
                Value = Properties.Settings.Default.CalculatorDecimalPlaces,
                Width = 60
            };
            decimalPlaces = Properties.Settings.Default.CalculatorDecimalPlaces;
            decimalPlacesInput.ValueChanged += (s, e) => decimalPlaces = (int)decimalPlacesInput.Value;

            mainPanel.Controls.Add(decimalLabel, 0, 0);
            mainPanel.Controls.Add(decimalPlacesInput, 1, 0);

            Button saveButton = new Button
            {
                Text = "Save",
                DialogResult = DialogResult.OK,
                Anchor = AnchorStyles.Right
            };
            
            // Save settings when clicking save
            saveButton.Click += (s, e) => 
            { 
                Properties.Settings.Default.CalculatorDecimalPlaces = decimalPlaces;
                Properties.Settings.Default.Save();
            };
            
            mainPanel.Controls.Add(saveButton, 1, 1);

            this.Controls.Add(mainPanel);
        }
    }

    public static class ExpressionEvaluator
    {
        public static double Evaluate(string expression, int decimalPlaces)
        {
            // Replace symbolic constants
            expression = expression.Replace("π", Math.PI.ToString());
            expression = expression.Replace("e", Math.E.ToString());

            // Handle power operator ^
            while (expression.Contains("^"))
            {
                int powerIndex = expression.IndexOf("^");
                int baseStartIndex = powerIndex - 1;
                while (baseStartIndex >= 0 && (char.IsDigit(expression[baseStartIndex]) || expression[baseStartIndex] == '.' || expression[baseStartIndex] == ')'))
                    baseStartIndex--;
                baseStartIndex++;
                int exponentEndIndex = powerIndex + 1;
                while (exponentEndIndex < expression.Length && (char.IsDigit(expression[exponentEndIndex]) || expression[exponentEndIndex] == '.' || expression[exponentEndIndex] == '('))
                    exponentEndIndex++;
                string baseStr = expression.Substring(baseStartIndex, powerIndex - baseStartIndex);
                string exponentStr = expression.Substring(powerIndex + 1, exponentEndIndex - powerIndex - 1);
                DataTable dt = new DataTable();
                double baseValue = Convert.ToDouble(dt.Compute(baseStr, ""));
                double exponentValue = Convert.ToDouble(dt.Compute(exponentStr, ""));
                double powerResult = Math.Pow(baseValue, exponentValue);
                string replacement = "(" + powerResult.ToString($"F{decimalPlaces}") + ")";
                expression = expression.Substring(0, baseStartIndex) + replacement + expression.Substring(exponentEndIndex);
            }

            expression = expression.Replace(" ", "").Replace("×", "*").Replace("÷", "/");

            // Evaluate functions recursively
            expression = EvaluateFunctions(expression, decimalPlaces);

            DataTable dtFinal = new DataTable();
            return Convert.ToDouble(dtFinal.Compute(expression, ""));
        }

        private static string EvaluateFunctions(string expression, int decimalPlaces)
        {
            string[] functions = { "sqrt", "sin", "cos", "tan", "log", "ln" };
            foreach (var func in functions)
            {
                while (expression.Contains(func + "("))
                {
                    int idx = expression.IndexOf(func + "(");
                    int openCount = 1;
                    int closeIdx = idx + func.Length + 1;
                    while (openCount > 0 && closeIdx < expression.Length)
                    {
                        if (expression[closeIdx] == '(') openCount++;
                        if (expression[closeIdx] == ')') openCount--;
                        closeIdx++;
                    }
                    if (openCount == 0)
                    {
                        string innerExpr = expression.Substring(idx + func.Length + 1, closeIdx - idx - func.Length - 2);
                        innerExpr = EvaluateFunctions(innerExpr, decimalPlaces);
                        DataTable dt = new DataTable();
                        double innerValue = Convert.ToDouble(dt.Compute(innerExpr, ""));
                        double result = 0;
                        switch (func)
                        {
                            case "sqrt": result = Math.Sqrt(innerValue); break;
                            case "sin": result = Math.Sin(innerValue); break;
                            case "cos": result = Math.Cos(innerValue); break;
                            case "tan": result = Math.Tan(innerValue); break;
                            case "log": result = Math.Log10(innerValue); break;
                            case "ln": result = Math.Log(innerValue); break;
                        }
                        string replacement = "(" + result.ToString($"F{decimalPlaces}") + ")";
                        expression = expression.Substring(0, idx) + replacement + expression.Substring(closeIdx);
                    }
                    else
                    {
                        break;
                    }
                }
            }
            return expression;
        }
    }
}
