using Python.Runtime;
using Newtonsoft.Json;

using System.Text.RegularExpressions;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Math;
using OfficeOpenXml.FormulaParsing.ExpressionGraph.CompileStrategy;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Text;


namespace MessageIngest
{
    public static class AnvilEngine
    {
        public static int entryIndex = 0;
        public static string emailFolderPath;
        
        public static void Process(string _emailFolderPath)
        {
            int count = 1;
            emailFolderPath = _emailFolderPath;
            InitializePythonEngine();

            int successfulParse = 0;
            int unsuccessfulParse = 0;

            string[] files = Directory.GetFiles(emailFolderPath);

            foreach (var file in files)
            {
                if (!file.StartsWith("_temp_") && file.EndsWith(".msg"))
                {
                    Console.WriteLine(">>>>>> Ingesting Message: " + count++ + " <<<<<<");
                    try
                    {
                        ProcessEmailFile(file, ref successfulParse, ref unsuccessfulParse);
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"[Process] Error processing file {file}: {ex.Message}");
                        unsuccessfulParse++;
                    }
                }
            }
        }

        private static void InitializePythonEngine()
        {
            var PythonPath = Helper.GetPythonRoot() + ";" + Helper.GetModulePath();
            Runtime.PythonDLL = Helper.GetDLLPath();
            PythonEngine.Initialize();
            PythonEngine.PythonPath = PythonPath;
        }

        private static void ProcessEmailFile(string file, ref int successfulParse, ref int unsuccessfulParse)
        {
            var (msg, py_msgpath) = ExtractMessage(file);

            TimeEntry timeEntry = AssignDataToDataModel(msg);

            var py_file = new PyString(file);
            var api_key = new PyString(GetAPIKey());
            var aliasList = new PyString(GetAliasList());
            var attachment_examples = new PyString(GetAttachmentExamples());

            var msgdataDict = ExtractClassificationData(py_file, api_key, attachment_examples);
            AssignClassificationDataToDataModel(timeEntry, msgdataDict);

            timeEntry.sentdate = ExtractSentDate(msg);

            timeEntry.narrative = GenerateNarrative(timeEntry, api_key);

            timeEntry.alias = GenerateClientMatter(timeEntry.subject, api_key, aliasList);

            ProcessClientMatter(timeEntry, ref successfulParse, ref unsuccessfulParse);

            FinalizeProcessing(timeEntry, file, msg);
        }

        private static (dynamic, PyString) ExtractMessage(string file)
        {
            dynamic extract_msg = Py.Import("extract_msg");
            dynamic sys = Py.Import("sys");
            sys.path.append(Path.Combine(@"G:\projects\MessageIngest\"));
            var temp_file = Helper.GetTempPath(file);
            var msg = extract_msg.openMsg(temp_file);
            var py_msgpath = new PyString(temp_file);
            return (msg, py_msgpath);
        }



        private static TimeEntry AssignDataToDataModel(dynamic msg)
        {
            return new TimeEntry
            {
                bcc = msg.bcc,
                cc = msg.cc,
                body = msg.body,
                alias = msg.defaultFolderName,
                filename = msg.filename,
                messageId = msg.messageId,
                sender = "SENDER OF EMAIL: " + msg.sender,
                subject = "SUBJECT OF EMAIL: " + msg.subject,
                recipients = "RECIPIENTS OF EMAIL: " + msg.to
            };
        }

        public static string GetMsgAttachments(PyString py_msgpath, PyString api_key, PyString attachment_examples)
        {
            // Define a list of image file extensions to check against
            List<string> imageExtensions = new List<string> { ".jpg", ".jpeg", ".png", ".gif", ".bmp", ".tiff", ".svg", ".webp" };

            // Invoke the Python method and return the result as a PyObject
            var Utility = Py.Import("utility");   
            PyObject pyList = Utility.InvokeMethod("get_attachments", new PyObject[] {py_msgpath});

            // Convert the PyObject (holding a Python list) to a C# list of strings
            List<string> msgAttachments = new List<string>();
            


            for (int i = 0; i < pyList.Length(); i++)
            {
                string attachment = pyList[i].ToString();
                string extension = Path.GetExtension(attachment).ToLower();

            
                // Skip adding the attachment if it is an image file
                if (!imageExtensions.Contains(extension))
                {
                    string simple = Utility.InvokeMethod("Simplify_Attachment", new PyObject[] { api_key, pyList[i],  attachment_examples }).ToString();
                    msgAttachments.Add(simple);
                }
            }

            // Console.WriteLine("Simplified Attachment Names: " + string.Join(", ", msgAttachments));
            return string.Join(", ", msgAttachments);
        }        
        private static PyDict ExtractClassificationData(PyString py_file, PyString api_key, PyString attachment_examples)
        {
            // Get the temporary file path from the helper
            var temp_file = Helper.GetTempPath(py_file.ToString());
            var py_temp_file = new PyString(temp_file);

            // Get message attachments based on the provided data
            string msgAttachments = GetMsgAttachments(py_temp_file, api_key, attachment_examples);

            // Import the utility module
            var Utility = Py.Import("utility");

            temp_file = Helper.GetTempPath(py_file.ToString());
            py_temp_file = new PyString(temp_file);
            // Fetch message data from the utility function
            var msgdata = Utility.InvokeMethod("get_msgdata", new PyObject[] { py_temp_file, api_key });

            // Convert the message data to a PyDict
            var msgdict = msgdata.As<PyDict>();

            // Add the attachment_examples to the PyDict
            
            if(msgAttachments.Length > 0)
            {
                msgdict["attachments"] = new PyString(msgAttachments);
                msgdict["hasattachments"] = new PyString("true");            
            }
            else
            {
                msgdict["attachments"] = new PyString("");
                msgdict["hasattachments"] = new PyString("false");   

            }
            
            // Return the modified PyDict
            // Console.WriteLine("[Dict]: "+ msgdict.ToString());
            return msgdict;
        }

        private static void AssignClassificationDataToDataModel(TimeEntry timeEntry, PyDict msgdataDict)
        {
            timeEntry.attachments = msgdataDict["attachments"].ToString();
            timeEntry.hasattachments = msgdataDict["hasattachments"].ToString();
            timeEntry.domain = msgdataDict["domain"].ToString();
            timeEntry.role = msgdataDict["role"].ToString();
        }

        private static string ExtractSentDate(dynamic msg)
        {
            string msgJson = msg.getJson();
            var values = JsonConvert.DeserializeObject<Dictionary<string, string>>(msgJson);
            return values["date"];
        }

        private static string GenerateNarrative(TimeEntry timeEntry, PyString api_key)
        {
            var Generate = Py.Import("generate");
            var subject = new PyString(timeEntry.subject);
            var sender = new PyString(timeEntry.sender);
            var recipient = new PyString(timeEntry.recipients);
            var body = new PyString(timeEntry.body);
            var attachments = new PyString(timeEntry.attachments);
            var Default_Examples = new PyString(GetNarrativeExamples(Helper.GetDefaultExampleFileName()));
            var SIZA_Examples = new PyString(GetNarrativeExamples(Helper.GetSIZAExampleFileName()));
            var SIA_Examples = new PyString(GetNarrativeExamples(Helper.GetSIAExampleFileName()));
            var SAE_Examples = new PyString(GetNarrativeExamples(Helper.GetSEAExampleFileName()));
            var REA_Examples = new PyString(GetNarrativeExamples(Helper.GetREAExampleFileName()));


                if (timeEntry.domain == "external")
                {
                    if(timeEntry.role == "sender" && timeEntry.hasattachments == "true" )   
                    {
                        Console.WriteLine("Invoking External Sender With Attachments Model");
                        return Generate.InvokeMethod("Narrative_SAE", new PyObject[] { api_key, recipient, sender, body, subject, attachments, SAE_Examples }).ToString();
                    }
                    else if(timeEntry.role == "sender" && timeEntry.hasattachments == "false" )
                    {
                        Console.WriteLine("Invoking External Sender Zero Attachments Model");
                        return Generate.InvokeMethod("Narrative_RAE", new PyObject[] { api_key, recipient, sender, body, subject, attachments, Default_Examples }).ToString();
                    }                   
                    else if(timeEntry.role == "recipient" && timeEntry.hasattachments == "true" )
                    {
                        Console.WriteLine("Invoking Recipient Attachments External");
                        return Generate.InvokeMethod("Narrative_RAE", new PyObject[] { api_key, recipient, sender, body, subject, attachments, REA_Examples }).ToString();
                    }
                    else
                    {
                        Console.WriteLine("Invoking External Default");
                        return Generate.InvokeMethod("Narrative_Default_External", new PyObject[] { api_key, recipient, sender, body, subject, Default_Examples}).ToString();
                    }
                                
                }
                else
                {
                    if(timeEntry.role == "sender" && timeEntry.hasattachments == "false")
                    {
                        Console.WriteLine("Invoking Sender Internal_Zero Attachments");
                        return Generate.InvokeMethod("Narrative_Internal_Zero_Attachments", new PyObject[] { api_key, recipient, sender, body, subject, SIZA_Examples }).ToString();
                    }
                    else if(timeEntry.role == "sender" && timeEntry.hasattachments == "true")
                    {
                        Console.WriteLine("Invoking Sender Internal Attachments");
                        return Generate.InvokeMethod("Narrative_Internal_Attachments", new PyObject[] { api_key, recipient, sender, body, subject, attachments, Default_Examples }).ToString();

                    }
                    else
                    {
                        Console.WriteLine("Invoking Sender Internal Attachments");
                        return Generate.InvokeMethod("Narrative_Internal_Attachments", new PyObject[] { api_key, recipient, sender, body, subject, attachments, Default_Examples }).ToString();
                    }

                 }

        }

        private static string GenerateClientMatter(string subject, PyString api_key, PyString aliasList)
        {
            var Generate = Py.Import("generate");
            var cmgenerated = Generate.InvokeMethod("ClientMatter", new PyObject[] { new PyString(subject), api_key, aliasList });
            return cmgenerated.ToString().Trim();
        }

        private static void ProcessClientMatter(TimeEntry timeEntry, ref int successfulParse, ref int unsuccessfulParse)
        {
            string clientstr = "";
            string matterstr = "";

            if (MatchCM(timeEntry.alias, out clientstr, out matterstr))
            {
                timeEntry.client = clientstr;
                timeEntry.matter = matterstr;
                successfulParse += 1;
            }
            else
            {
                timeEntry.client = clientstr;
                timeEntry.matter = matterstr;
                unsuccessfulParse += 1;
            }
            // Console.WriteLine("[ProcessClientMatter]: " + clientstr + "-" + matterstr);
        }

        private static void FinalizeProcessing(TimeEntry timeEntry, string file, dynamic msg)
        {
            string msgdate = msg.date.ToString();
            timeEntry.date = Helper.ConvertDate(msgdate);

            AppendEmailManifest(file);
            AppendJson(timeEntry);
            Helper.AddRowXL(timeEntry);

            Console.WriteLine("Narrative: " + timeEntry.narrative);
            
        }


    
        

        public static void AppendJson(TimeEntry entry)
        {
            // Console.WriteLine(entry.subject);
            TimeEntry[] timeEntries = new TimeEntry[1000];
            string jsonFilename = Helper.GetJsonPath();
            try{
                
                 timeEntries = JsonConvert.DeserializeObject<TimeEntry[]>(File.ReadAllText(jsonFilename));
            }
            catch { }
            

                timeEntries[entryIndex++] = entry;
                
               

                
            
            

            
           
            // serialize JSON to a string and then write string to a file
            File.WriteAllText(jsonFilename, JsonConvert.SerializeObject(timeEntries, Formatting.Indented));
        }

        public static void AppendEmailManifest(string name) 
        {

            string path = emailFolderPath + "\\" + "manifest.txt";
            File.AppendAllLines(path, new [] { name });
            // Console.WriteLine(name);
        }

        public static bool CheckJson(TimeEntry entry, TimeEntry[] timeEntries)  
        {
            foreach (var e in timeEntries)
            {
                if (e.messageId == entry.messageId)
                {
                    return false;
                }
            }

            return true;

        }
        
        public static string GetAliasList()
        {
            List<ClientMatter> CMList  = ExcelReader.ReadClientMatterExcel(Helper.GetClientDataFileName());
            string aliaslist = "None;";
            foreach (ClientMatter cm in CMList)
            {
                if(cm.Partner != null  && cm.Partner.ToLower().Contains("cravens"))
                {
                    string removed = cm.Alias != null ? CleanAlias(cm.Alias.ToLower()) : string.Empty;   
                    aliaslist += removed + ";";
                }

                                 
            }
            return aliaslist;
            
        }

        public static string RemoveCommas(string rawline)
        {
                string pattern = "(?<=\\\"[^\\\"]*),(?=[^\\\"]*\\\")";
                rawline = Regex.Replace(rawline, pattern, "");
                return rawline;
        }

        public static string CleanAlias(string rawline)
        {
            
            rawline = RemoveEscapedQuotes(rawline);
           
            return rawline;

        }
        public static string RemoveEscapedQuotes(string rawline)
        { 
                string pattern = @"\?s";
                rawline = Regex.Replace(rawline, pattern, "s");
                // Console.WriteLine("[RemoveEscapedQuotes]: " + rawline);
                return rawline;
               
        }

         public static string GetNarrativeExamples(string filename)
        {
            string examples = "";
            foreach (string line in File.ReadLines(filename))
            {
                examples += line;
                                 
            }
            return examples;
            
        }
        public static string GetAttachmentExamples()
        {
            string examples = "";
            foreach (string line in File.ReadLines(Helper.GetAttachmentExampleFileName()))
            {
                examples += line;
                                 
            }
            return examples;
            
        }



        public static string GetAPIKey()
        {
            var dict = File.ReadLines(@"G:\secret.txt").Select(line => line.Split(',')).ToDictionary(line => line[0], line => line[1]);
            
            if(dict.TryGetValue("API_Key", out var apikey))
            {
                return apikey;
            }

            return "";
    

        }

      
        
        public static bool MatchCM(string generatedAlias, out string clientstr, out string matterstr)
        {
                       
            List<ClientMatter> CMList =  ExcelReader.ReadClientMatterExcel(Helper.GetClientDataFileName());
            int i = 0;
            double similarity = 0;
            int result = -1;
            foreach(ClientMatter reference in CMList)
            {
                double newSimilarity = StringSimilarity.CalculateSimilarity(reference.Alias, generatedAlias);
                if(similarity < newSimilarity)
                {
                    if(newSimilarity == 1) 
                    {
                        similarity = newSimilarity;
                        result = i ;
                        break;
                    }
                    else
                    {
                        similarity = newSimilarity;
                        result = i;
                    }
                    
                }
                i++;
            }
            

            

            clientstr = CMList[result].Client;
            matterstr = CMList[result].Matter;

            if(similarity >= .95)
            {
                return true;
            }
            else
            {
                return false;
            }
        }
            
        
        public static void Shutdown()
        {
            try {
                PythonEngine.Shutdown();
            }
            catch {}
        }
   

        
    }
}