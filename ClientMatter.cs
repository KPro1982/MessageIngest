using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using OfficeOpenXml;

namespace MessageIngest
{
    public class ClientMatter
    {
        public string Alias { get; set; }
        public string CMString { get; set; }
        public string Client { get; set; }

        public string Matter { get; set; }  
        public string Partner { get; set; }
        // Add other properties as needed
    }

    public static class ExcelReader
    {
        public static List<ClientMatter> ReadClientMatterExcel(string filePath)
        {
            var clientMatters = new List<ClientMatter>();

            // Ensure the file exists
            if (!File.Exists(filePath))
                throw new FileNotFoundException("File not found", filePath);

            // Load the Excel file
            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                // Get the first worksheet

                var worksheet = package.Workbook.Worksheets.First<ExcelWorksheet>();

                // Start reading from row 2 (assuming row 1 contains headers)

                int startRow = 2;
                int endRow = worksheet.Dimension.End.Row;

                for (int row = startRow; row <= endRow; row++)
                {
                    var clientMatter = new ClientMatter
                    {
                        Alias = worksheet.Cells[row, 1].Value?.ToString(),
                        CMString = worksheet.Cells[row, 2].Value?.ToString(),
                        Partner = worksheet.Cells[row, 3].Value?.ToString()
                    };
                    string client, matter;
                    if(clientMatter.CMString != null)
                    {
                        ParseClientMatter(clientMatter.CMString, out client, out matter);
                        clientMatter.Client = client;
                        clientMatter.Matter = matter;
                    }
                    else
                    {
                        clientMatter.Client = "0000";
                        clientMatter.Matter = "00000";

                    }
                    

                    clientMatters.Add(clientMatter);
                }
            }

            return clientMatters;
        }
    


        public static void ParseClientMatter(string CmStr, out string Client, out string Matter)
        {
            string[] parts = CmStr.Split(new char[] { '.', '-' }, 2);
            
            // Ensure there are two parts after the split
            if (parts.Length == 2)
            {
                Client = parts[0];
                Matter = parts[1];
                
                // Console.WriteLine("[SplitClientMatter] Client: " + Client);
                // Console.WriteLine("[SplitClientMatter] Matter: " + Matter);
            }
            else
            {
                Client = "0000";
                Matter = "00000";
                Console.WriteLine("Error [SplitClientMatter]: CMString does not contain a '.' or '-'.");
            }
        }
    }
}
