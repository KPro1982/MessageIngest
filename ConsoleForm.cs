using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System;
using System.IO;
using System.Windows.Forms;
using System.Text;

namespace MessageIngest
{
   using System;
using System.Text;
using System.Windows.Forms;

    public class ConsoleOutputForm : Form
    {
        private TextBox outputTextBox;
        private Button closeButton;

        public ConsoleOutputForm()
        {
            // Set the form properties
            this.Text = "Console Output";
            this.Width = 2500;
            this.Height = 1200;
            

            // Create and configure the TextBox
            outputTextBox = new TextBox
            {
                Multiline = true,
                Dock = DockStyle.Fill,
                ScrollBars = ScrollBars.Vertical,
                ReadOnly = true
            };

            // Create and configure the Close button
            closeButton = new Button
            {
                Text = "Close",
                Dock = DockStyle.Bottom,
                Height = 80
            };
            closeButton.Click += CloseButton_Click;

            // Add the TextBox and Close button to the form
            this.Controls.Add(outputTextBox);
            this.Controls.Add(closeButton);

            // Redirect the console output to the TextBox
            var writer = new TextBoxWriter(outputTextBox);
            Console.SetOut(writer);
        }

        private void CloseButton_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        // A custom TextWriter class to write to the TextBox
        private class TextBoxWriter : TextWriter
        {
            private TextBox _textBox;

            public TextBoxWriter(TextBox textBox)
            {
                _textBox = textBox;
            }

            public override void Write(char value)
            {
                _textBox.Invoke((Action)(() => _textBox.AppendText(value.ToString())));
            }

            public override void Write(string value)
            {
                _textBox.Invoke((Action)(() => _textBox.AppendText(value)));
            }

            public override Encoding Encoding
            {
                get { return System.Text.Encoding.UTF8; }
            }
        }
    }

}