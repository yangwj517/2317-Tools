using System.Windows;

namespace ToolboxHost
{
    public partial class InputDialog : Window
    {
        public string Answer { get; set; }

        public InputDialog(string title, string prompt, string defaultAnswer = "")
        {
            InitializeComponent();
            this.Title = title;
            this.PromptText.Text = prompt;
            this.Answer = defaultAnswer;
            this.DataContext = this;
        }

        private void OKButton_Click(object sender, RoutedEventArgs e)
        {
            this.DialogResult = true;
        }
    }
}