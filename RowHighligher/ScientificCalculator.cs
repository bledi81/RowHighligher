using System;
using System.Drawing;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;
using System.Data;

namespace RowHighligher
{
    public partial class ScientificCalculator : Form
    {
        private TextBox displayTextBox;
        private Button insertButton;
        private Button clearButton;
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

        public ScientificCalculator()
        {
            InitializeComponents();

            // Set form properties
            this.FormBorderStyle = FormBorderStyle.SizableToolWindow;
            this.Text = "Scientific Calculator";
            this.MinimumSize = new Size(300, 400);
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
                RowCount = 8, // Increase from 7 to 8
                ColumnCount = 1,
                Padding = new Padding(10)
            };
            mainPanel.RowStyles.Add(new RowStyle(SizeType.Percent, 15)); // Display
            mainPanel.RowStyles.Add(new RowStyle(SizeType.Percent, 10)); // Insert/Clear
            mainPanel.RowStyles.Add(new RowStyle(SizeType.Percent, 10)); // Parentheses (new)
            mainPanel.RowStyles.Add(new RowStyle(SizeType.Percent, 10)); // Scientific
            mainPanel.RowStyles.Add(new RowStyle(SizeType.Percent, 15)); // Numbers
            mainPanel.RowStyles.Add(new RowStyle(SizeType.Percent, 15)); // Numbers
            mainPanel.RowStyles.Add(new RowStyle(SizeType.Percent, 15)); // Numbers
            mainPanel.RowStyles.Add(new RowStyle(SizeType.Percent, 10)); // Bottom row

            // Display textbox
            displayTextBox = new TextBox
            {
                Dock = DockStyle.Fill,
                ReadOnly = true,
                Font = new Font("Consolas", 16, FontStyle.Bold),
                TextAlign = HorizontalAlignment.Right,
                Text = "0"
            };
            mainPanel.Controls.Add(displayTextBox, 0, 0);

            // Insert button in a panel with Clear
            TableLayoutPanel buttonsPanel = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 2,
                RowCount = 1
            };
            buttonsPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50));
            buttonsPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50));

            insertButton = new Button
            {
                Text = "Insert (Ctrl+Enter)",
                Dock = DockStyle.Fill
            };
            insertButton.Click += InsertButton_Click;
            buttonsPanel.Controls.Add(insertButton, 0, 0);

            clearButton = new Button
            {
                Text = "LastAns",
                Dock = DockStyle.Fill
            };
            clearButton.Click += LastAnsButton_Click; // Change the click handler
            buttonsPanel.Controls.Add(clearButton, 1, 0);

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



            // Scientific buttons panel
            TableLayoutPanel scientificPanel = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 5,
                RowCount = 1
            };
            scientificPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 20));
            scientificPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 20));
            scientificPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 20));
            scientificPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 20));
            scientificPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 20));

            scientificButtons = new Button[5];
            string[] scientificOps = { "sqrt", "x^y", "1/x", "sin", "cos" };

            for (int i = 0; i < scientificButtons.Length; i++)
            {
                scientificButtons[i] = new Button
                {
                    Text = scientificOps[i],
                    Dock = DockStyle.Fill,
                    Tag = scientificOps[i]
                };
                scientificButtons[i].Click += ScientificButton_Click;
                scientificPanel.Controls.Add(scientificButtons[i], i, 0);
            }

            mainPanel.Controls.Add(scientificPanel, 0, 2);

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

            mainPanel.Controls.Add(numberPanel, 0, 3);
            mainPanel.SetRowSpan(numberPanel, 3);

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

            mainPanel.Controls.Add(bottomPanel, 0, 6);

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
                    keyBuffer.EndsWith("cos", StringComparison.OrdinalIgnoreCase))
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
                    
                    // Remove the function characters and add the function with parenthesis
                    string currentText = displayTextBox.Text;
                    displayTextBox.Text = currentText.Substring(0, currentText.Length - charsToCut) + 
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
            // Handle decimal point
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
            else if (!displayTextBox.Text.Contains("."))
            {
                displayTextBox.Text += ".";
            }
            FormatDisplayText(); // Call to format the display text for better readability
            displayTextBox.SelectionStart = displayTextBox.Text.Length; // Move cursor to the end
            displayTextBox.SelectionLength = 0; // Clear selection
        }

        private void OperatorButton_Click(object sender, EventArgs e)
        {
            Button button = (Button)sender;
            HandleOperation(button.Text);
        }

        private void HandleOperation(string operation)
        {
            if (isExpressionMode)
            {
                // In expression mode, just append the operator but with proper spacing
                displayTextBox.Text += " " + operation + " ";
                
                // Important: Don't call FormatDisplayText() here which would interfere with functions
                // but we do want to ensure the operator is visible
                
                // Always set cursor to end
                SetCursorToEnd();
            }
            else if (!isNewCalculation)
            {
                // Your existing code...
                if (lastOperation != "")
                {
                    CalculateResult();
                }

                lastValue = double.Parse(displayTextBox.Text);
                lastOperation = operation;

                // Show the operation in the display
                displayTextBox.Text += " " + operation + " ";
                isNewCalculation = false;
            }
            else
            {
                // Your existing code...
            }
        }

        private void EqualsButton_Click(object sender, EventArgs e)
        {
            CalculateResult();
        }

        private void CalculateResult()
        {
            try
            {
                if (isNewCalculation || displayTextBox.Text == "0")
                {
                    return; // No calculation needed
                }

                // Single check for balanced parentheses using the more reliable method
                if (!HasBalancedParentheses(displayTextBox.Text))
                {
                    // Only show error for actual expressions that contain parentheses
                    if (displayTextBox.Text.Contains("(") || displayTextBox.Text.Contains(")"))
                    {
                        MessageBox.Show("Expression has unbalanced parentheses. Please close all open brackets.",
                            "Syntax Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        // Reset the bracket count to match reality
                        bracketCount = CountOpenParentheses(displayTextBox.Text);
                        return;
                    }
                }



                // Reset expression mode and bracket count if parentheses are balanced
                isExpressionMode = false;
                bracketCount = 0;

                // Prepare the expression for calculation - use our improved function
                string expression = PrepareExpressionForEvaluation(displayTextBox.Text);

                // Debug the expression to see what we're calculating
                // Uncomment this line when testing
                // DebugExpression(expression);

                // Use DataTable to evaluate the expression
                DataTable dt = new DataTable();
                var result = dt.Compute(expression, "");

                // Process and display the result
                if (result is DBNull)
                {
                    displayTextBox.Text = "Error";
                }
                else
                {
                    double calculatedValue = Convert.ToDouble(result);
                    displayTextBox.Text = calculatedValue.ToString();
                    lastAnswer = calculatedValue; // Store the last answer
                }

                // Reset state
                lastOperation = "";
                isNewCalculation = true;
                isExpressionMode = false;
                bracketCount = 0;

                SetCursorToEnd();

            }
            catch (Exception ex)
            {
                displayTextBox.Text = "Error";
                isNewCalculation = true;
                isExpressionMode = false;
                MessageBox.Show($"Error calculating result: {ex.Message}", "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
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
            // First, handle any ^ (power) operations
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
                string replacement = "(" + powerResult.ToString() + ")";
                expression = expression.Substring(0, baseStartIndex) + 
                             replacement + 
                             expression.Substring(exponentEndIndex);
            }

            expression = expression.Replace(" ", "")
                                  .Replace("×", "*")
                                  .Replace("÷", "/");

            try {
                // Process functions one at a time
                while (expression.Contains("sqrt(") || expression.Contains("sin(") || expression.Contains("cos("))
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
                            string replacement = "(" + sqrtValue.ToString() + ")";
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
                            string replacement = "(" + sinValue.ToString() + ")";
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
                            
                            string replacement = "(" + cosValue.ToString() + ")";
                            expression = expression.Substring(0, cosIdx) + replacement + expression.Substring(closeIdx);
                            
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

            // Handle functions that require parentheses
            if (operation == "sqrt" || operation == "sin" || operation == "cos")
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
            lastValue = 0;
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
    }
}
