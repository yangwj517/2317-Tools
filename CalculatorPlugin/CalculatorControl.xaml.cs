using System;
using System.Windows;
using System.Windows.Controls;

namespace CalculatorPlugin
{
    public partial class CalculatorControl : UserControl
    {
        private string currentInput = "0";
        private string previousInput = "";
        private char currentOperator = '\0';
        private bool resetOnNextInput = false;

        public CalculatorControl()
        {
            InitializeComponent();
            UpdateDisplay();
        }

        private void UpdateDisplay()
        {
            DisplayTextBox.Text = currentInput;
            HistoryTextBlock.Text = previousInput + (currentOperator != '\0' ? " " + currentOperator : "");
        }

        private void NumberButton_Click(object sender, RoutedEventArgs e)
        {
            if (sender is Button button)
            {
                string number = button.Content.ToString();

                if (currentInput == "0" || resetOnNextInput)
                {
                    currentInput = number;
                    resetOnNextInput = false;
                }
                else
                {
                    currentInput += number;
                }

                UpdateDisplay();
            }
        }

        private void OperatorButton_Click(object sender, RoutedEventArgs e)
        {
            if (sender is Button button)
            {
                char newOperator = button.Content.ToString()[0];

                if (currentOperator != '\0' && !resetOnNextInput)
                {
                    // 如果已经有操作符，先计算之前的结果
                    CalculateResult();
                }

                previousInput = currentInput;
                currentOperator = newOperator;
                resetOnNextInput = true;
                UpdateDisplay();
            }
        }

        private void EqualsButton_Click(object sender, RoutedEventArgs e)
        {
            CalculateResult();
            currentOperator = '\0';
            UpdateDisplay();
        }

        private void CalculateResult()
        {
            if (string.IsNullOrEmpty(previousInput) || currentOperator == '\0')
                return;

            try
            {
                double prev = double.Parse(previousInput);
                double current = double.Parse(currentInput);
                double result = 0;

                switch (currentOperator)
                {
                    case '+': result = prev + current; break;
                    case '-': result = prev - current; break;
                    case '×': result = prev * current; break;
                    case '÷':
                        if (current == 0)
                        {
                            currentInput = "错误";
                            resetOnNextInput = true;
                            UpdateDisplay();
                            return;
                        }
                        result = prev / current;
                        break;
                }

                currentInput = result.ToString();
                previousInput = "";
                resetOnNextInput = true;
            }
            catch (Exception ex)
            {
                currentInput = "错误";
                resetOnNextInput = true;
            }
        }

        private void ClearButton_Click(object sender, RoutedEventArgs e)
        {
            currentInput = "0";
            previousInput = "";
            currentOperator = '\0';
            resetOnNextInput = false;
            UpdateDisplay();
        }

        private void DecimalButton_Click(object sender, RoutedEventArgs e)
        {
            if (!currentInput.Contains("."))
            {
                currentInput += ".";
                UpdateDisplay();
            }
        }

        private void SignButton_Click(object sender, RoutedEventArgs e)
        {
            if (currentInput != "0")
            {
                if (currentInput.StartsWith("-"))
                {
                    currentInput = currentInput.Substring(1);
                }
                else
                {
                    currentInput = "-" + currentInput;
                }
                UpdateDisplay();
            }
        }

        private void PercentButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                double value = double.Parse(currentInput);
                currentInput = (value / 100).ToString();
                UpdateDisplay();
            }
            catch
            {
                currentInput = "错误";
                resetOnNextInput = true;
                UpdateDisplay();
            }
        }
    }
}