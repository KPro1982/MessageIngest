using System.CodeDom;
using System.Windows.Forms;
using System.Text;

using Accessibility;

namespace MessageIngest
{
    
    static class Program
    {
        static string _folderpath;
        
        [STAThread]
        static void Main()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
            ApplicationConfiguration.Initialize();
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            // Get the folder path before showing the non-modal console form
            _folderpath = GetEmailFolderPath();

            if (string.IsNullOrEmpty(_folderpath))
            {
                MessageBox.Show("No folder selected. Exiting application.");
                return;
            }
            // Create and show the console output form as a non-modal window
            var consoleOutputForm = new ConsoleOutputForm();

            consoleOutputForm.FormClosed += (sender, e) => Application.Exit();

            // Get the folder path before starting the task

            consoleOutputForm.Show();
            

            Task.Run(() =>
            {
                Console.WriteLine("Message Ingestion Running...");
                Helper.SetEmailFolderPath(_folderpath);
                AnvilEngine.Process(_folderpath);
                AnvilEngine.Shutdown();
                Helper.DeleteTempFiles(_folderpath);  
            });

            Application.Run();    
        }


        public static string GetEmailFolderPath()
        {

            using (FolderBrowserDialog folderDlg = new FolderBrowserDialog())
            {
                DialogResult result = folderDlg.ShowDialog();
                if (result == DialogResult.OK)
                {
                    _folderpath = folderDlg.SelectedPath;
                }
            }
            

            return _folderpath;
        }
    }
}