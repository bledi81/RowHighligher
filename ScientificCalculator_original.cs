using System;
using System.Drawing;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

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
        private bool isInKeyboardMode = false;

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

        private void InitializeComponents()
        {
            // Main layout panel
            TableLayoutPanel mainPanel = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                RowCount = 7,
                ColumnCount = 1,
                Padding = new Padding(10)
            };
            mainPanel.RowStyles.Add(new RowStyle(SizeType.Percent, 15));
            mainPanel.RowStyles.Add(new RowStyle(SizeType.Percent, 10));
            mainPanel.RowStyles.Add(new RowStyle(SizeType.Percent, 15));
            mainPanel.RowStyles.Add(new RowStyle(SizeType.Percent, 15));
            mainPanel.RowStyles.Add(new RowStyle(SizeType.Percent, 15));
            mainPanel.RowStyles.Add(new RowStyle(SizeType.Percent, 15));
            mainPanel.RowStyles.Add(new RowStyle(SizeType.Percent, 15));

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
                Text = "Clear",
                Dock = DockStyle.Fill
            };
            clearButton.Click += ClearButton_Click;
            buttonsPanel.Controls.Add(clearButton, 1, 0);

            mainPanel.Controls.Add(buttonsPanel, 0, 1);

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
            string[] scientificOps = { "sqrt", "x²", "1/x", "sin", "cos" };

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

        private void ScientificCalculator_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Control && e.KeyCode == Keys.Enter)
            {
                InsertButton_Click(sender, e);
                e.Handled = true;
            }
        }

        private void ScientificCalculator_KeyPress(object sender, KeyPressEventArgs e)
        {
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
                displayTextBox.Text = value;
                isNewCalculation = false;
            }
            else
            {
                if (displayTextBox.Text == "0")
                    displayTextBox.Text = value;
                else
                    displayTextBox.Text += value;
            }
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
        }

        private void OperatorButton_Click(object sender, EventArgs e)
        {
            Button button = (Button)sender;
            HandleOperation(button.Text);
        }

        private void HandleOperation(string operation)
        {
            if (!isNewCalculation)
            {
                if (lastOperation != "")
                {
                    CalculateResult();
                }
                
                lastValue = double.Parse(displayTextBox.Text);
                lastOperation = operation;
                isNewCalculation = true;
            }
            else
            {
                // Just change the operation if one was already entered
                lastOperation = operation;
            }
        }

        private void EqualsButton_Click(object sender, EventArgs e)
        {
            CalculateResult();
        }

        private void CalculateResult()
        {
            if (lastOperation != "")
            {
                double currentValue = double.Parse(displayTextBox.Text);
                double result = 0;

                switch (lastOperation)
                {
                    case "+":
                        result = lastValue + currentValue;
                        break;
                    case "-":
                        result = lastValue - currentValue;
                        break;
                    case "*":
                        result = lastValue * currentValue;
                        break;
                    case "/":
                        if (currentValue != 0)
                            result = lastValue / currentValue;
                        else
                        {
                            MessageBox.Show("Cannot divide by zero", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            return;
                        }
                        break;
                }

                displayTextBox.Text = result.ToString();
                lastOperation = "";
                isNewCalculation = true;
            }
        }

        private void ScientificButton_Click(object sender, EventArgs e)
        {
            Button button = (Button)sender;
            string operation = button.Tag.ToString();
            double value;

            if (double.TryParse(displayTextBox.Text, out value))
            {
                double result = 0;

                switch (operation)
                {
                    case "sqrt":
                        if (value >= 0)
                            result = Math.Sqrt(value);
                        else
                        {
                            MessageBox.Show("Cannot calculate square root of negative value", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            return;
                        }
                        break;
                    case "x²":
                        result = value * value;
                        break;
                    case "1/x":
                        if (value != 0)
                            result = 1 / value;
                        else
                        {
                            MessageBox.Show("Cannot divide by zero", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            return;
                        }
                        break;
                    case "sin":
                        result = Math.Sin(value);
                        break;
                    case "cos":
                        result = Math.Cos(value);
                        break;
                }

                displayTextBox.Text = result.ToString();
                isNewCalculation = true;
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
    }
}
