using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Data.OleDb;
using System.Data;
using Python.Runtime;
using OfficeOpenXml;
using System.Diagnostics;
using System.Text.RegularExpressions;



namespace MessageIngest
{
    public static class Helper
    {
        public static string _folderpath = "";
        public static void SetEmailFolderPath(string folderPath)
        {
            _folderpath = folderPath;
        }
        public static string GetTempPath(string msg_path)
        {
            string temp_file = _folderpath  + @"\" + "_temp_" + Path.GetRandomFileName() + ".msg";
            // Console.WriteLine("Temp Path: " + temp_file);    
            File.Copy(msg_path, temp_file);
            return temp_file;

        }

        public static void CloseTempPath(string temp_file)
        {
            File.Delete(temp_file);
        }

        public static void DeleteTempFiles(string folderPath)
        {

            TerminateOutlookProcesses();
            try
            {
                TerminateOutlookProcesses();
                // Get all files in the specified folder
                string[] files = Directory.GetFiles(folderPath);

                // Loop through the files
                foreach (string file in files)
                {
                    // Get the file name
                    string fileName = Path.GetFileName(file);

                    // Check if the file name starts with "_temp_"
                    if (fileName.StartsWith("_temp_", StringComparison.OrdinalIgnoreCase))
                    {
                        // Delete the file
                        File.Delete(file);
                        // Console.WriteLine($"Deleted: {file}");
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[DeleteTempFiles]: {ex.Message}");
            }
        }
        static void TerminateOutlookProcesses()
        {
            // Get all processes that have the name "OUTLOOK"
            Process[] outlookProcesses = Process.GetProcessesByName("OUTLOOK");

            if (outlookProcesses.Length > 0)
            {
                foreach (Process process in outlookProcesses)
                {
                    try
                    {
                        // Kill the process
                        process.Kill();
                        Console.WriteLine("Outlook process terminated: " + process.Id);
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Failed to terminate Outlook process {process.Id}: {ex.Message}");
                    }
                }
            }
            else
            {
                Console.WriteLine("No Outlook processes are running.");
            }
        }
        public static string GetDLLPath()
        {
            return Environment.ExpandEnvironmentVariables(@"%LOCALAPPDATA%\Programs\Python\Python312\python312.dll");

        }
        public static string GetPythonRoot()
        {
            return Environment.ExpandEnvironmentVariables(@"%LOCALAPPDATA%\Programs\Python\Python312");
        }

        public static string GetModulePath()
        {
            return Environment.ExpandEnvironmentVariables(@"%LOCALAPPDATA%\Programs\Python\Python312\Lib\site-packages");
        }

        public static string GetExecutionPath()
        {
            string? currentPath = Path.GetDirectoryName(System.AppContext.BaseDirectory);
            if (currentPath != null) {
                return  Path.GetFullPath(Path.Combine(currentPath, "."));
            }
            else{
                return "";
            }
            

        }
        public static string GetJsonPath()
        {
            return @"G:\timeentries.json";
        }

        public static string GetClientDataFileName()
        {
            return @"G:\ClientMatter.xlsx";
        }
        public static string GetDefaultExampleFileName()
        {
            return @"G:\default_timeexamples.csv";
        }
        public static string GetSIZAExampleFileName()
        {
            return @"G:\Sent_Internal_Zero_Attachments_TimeExamples.csv";
        }

        public static string GetSIAExampleFileName()
        {
            return @"G:\Sent_Internal_Attachments_TimeExamples.csv";
        }
        public static string GetSEAExampleFileName()
        {
            return @"G:\Sent_External_Attachments_TimeExamples.csv";
        }
        public static string GetREAExampleFileName()
        {
            return @"G:\Received_External_Attachments_TimeExamples.csv";
        }




        public static string GetAttachmentExampleFileName()
        {
            return @"G:\attachmentexamples.csv";
        }
        
        
        

        public static string ConvertDate(string msgdate)
        {
            DateTime dt = DateTime.Parse(msgdate);
            string dtstr = dt.ToString("yyyyMMdd");
            

            return dtstr;
;
        }
        public static string ConvertTime(string msgtime)
        {
            DateTime dt = DateTime.Parse(msgtime);
            string dtstr = dt.ToString("H:mm:ss");
            

            return dtstr;
;
        }

       public static void CreateXL(string path = "")
        {
            if(path == "")
            {
                path = Helper.GetXLPath();
            }

            FileInfo fileInfo = new FileInfo(path);

            if (!fileInfo.Exists)
            {
                using (ExcelPackage package = new ExcelPackage(fileInfo))
                {
                    // Add a worksheet to the empty workbook if you want
                    ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("Sheet1");

                    worksheet.Cells[1, 1].Value = "UserId";
                    worksheet.Cells[1, 2].Value = "Date";
                    worksheet.Cells[1, 3].Value = "Timekeeper";
                    worksheet.Cells[1, 4].Value = "Client";
                    worksheet.Cells[1, 5].Value = "Matter";
                    worksheet.Cells[1, 6].Value = "Task";                    
                    worksheet.Cells[1, 7].Value = "client";
                    worksheet.Cells[1, 7].Value = "Activity";
                    worksheet.Cells[1, 8].Value = "Billable";
                    worksheet.Cells[1, 9].Value = "HoursWorked";
                    worksheet.Cells[1, 10].Value = "HoursBilled";
                    worksheet.Cells[1, 11].Value = "Rate";
                    worksheet.Cells[1, 12].Value = "Amount";
                    worksheet.Cells[1, 13].Value = "Code1";
                    worksheet.Cells[1, 14].Value = "Code2";
                    worksheet.Cells[1, 15].Value = "Code3";
                    worksheet.Cells[1, 16].Value = "Note";
                    worksheet.Cells[1, 17].Value = "Time";
                    worksheet.Cells[1, 18].Value = "Narrative";
                    worksheet.Cells[1, 19].Value = "Body";
                    worksheet.Cells[1, 20].Value = "Subject";
                    worksheet.Cells[1, 21].Value = "Sentdate";
                    worksheet.Cells[1, 22].Value = "Attachments";
                    worksheet.Cells[1, 23].Value = "HasAttachments";
                    worksheet.Cells[1, 24].Value = "Domain";
                    worksheet.Cells[1, 25].Value = "Role";




                    package.Save();
                }
                // System.Console.WriteLine("Excel file created at: " + path);
            }
            else
            {
               //  System.Console.WriteLine("Excel file already exists at: " + path);
            }
        }

        public static string GetXLPath()
        {

            return _folderpath  + "\\" + "cravens_timeentry.xlsx";
        }
        public static string RemoveExtraLineBreaks(string input)
        {
            // Use Regex to replace multiple consecutive line breaks (with or without spaces/tabs between them) with a single line break
            return Regex.Replace(input, @"(\r?\n\s*){2,}", "\n");
        }
        public static void AddRowXL(TimeEntry data, string path = "")
        {
            if(path == "")
            {
                path = Helper.GetXLPath();
            }

            FileInfo fileInfo = new FileInfo(path);

            if (!fileInfo.Exists)
            {
                CreateXL(path);
                fileInfo = new FileInfo(path);
            }    

            if (fileInfo.Exists)
            {
                using (ExcelPackage package = new ExcelPackage(fileInfo))
                {
                    // Get the first worksheet
                    ExcelWorksheet worksheet = package.Workbook.Worksheets.First<ExcelWorksheet>();

                    // Find the next empty row
                    int newRow = worksheet.Dimension.End.Row + 1;

                    // Add the data to the new row

                    worksheet.Cells[newRow, 1].Value = data.userId;
                    worksheet.Cells[newRow, 2].Value = data.date;
                    worksheet.Cells[newRow, 3].Value = data.timekeepr; 
                    worksheet.Cells[newRow, 4].Value = data.client;
                    worksheet.Cells[newRow, 5].Value = data.matter; 
                    worksheet.Cells[newRow, 6].Value = data.task;
                    worksheet.Cells[newRow, 7].Value = data.activity;
                    worksheet.Cells[newRow, 8].Value = data.billable; 
                    worksheet.Cells[newRow, 9].Value = data.hoursWorked;
                    worksheet.Cells[newRow, 10].Value = data.hoursBilled;
                    worksheet.Cells[newRow, 11].Value = data.rate;
                    worksheet.Cells[newRow, 12].Value = data.amount;
                    // worksheet.Cells[newRow, 16].Value = "Note";
                    // worksheet.Cells[newRow, 17].Value = "Time";
                    worksheet.Cells[newRow, 18].Value = data.narrative;
                    worksheet.Cells[newRow, 19].Value = RemoveExtraLineBreaks(data.body); 
                    worksheet.Cells[newRow, 25].Value = data.alias; 
                    worksheet.Cells[newRow, 21].Value = data.sentdate;
                    worksheet.Cells[newRow, 22].Value = data.attachments; 
                    worksheet.Cells[newRow, 23].Value = data.hasattachments; 
                    worksheet.Cells[newRow, 24].Value = data.domain; 
                    worksheet.Cells[newRow, 20].Value = data.role; 
                    worksheet.Cells[newRow,26].Value =  data.subject;    
                    

                    // Save the changes
                    package.Save();
                }
                //System.Console.WriteLine($"Added new row for user {data.userId} to Excel file at: " + path);
            }
            else
            {
                //System.Console.WriteLine("The Excel file does not exist at the specified path: " + path);
            }
        }

    }
}