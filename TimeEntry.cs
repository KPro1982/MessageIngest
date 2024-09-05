using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace MessageIngest
{
    public class TimeEntry
    {
        public TimeEntry()
        {
            this.userId = "DCRA";
            this.timekeepr = "DCRA";
            this.sentdate = "";
            this.billable = "y";
        }

        
        public string? userId { get; set; }
         public string? timekeepr { get; set; }
        public string? client { get; set; }       
        public string? matter { get; set; }
         public string? task { get; set; }
        public string? activity { get; set; }       
        public string? billable { get; set; }           
        public string? hoursWorked { get; set; }           
        public string? hoursBilled { get; set; }           
        public string? rate { get; set; }   
        public string? amount { get; set; }   
        public string? narrative  { get; set; }   
        public string? alias { get; set; }   
        public int? length { get; set; } 
        public string? subject { get; set; }
        public string? bcc { get; set; }
        public string? body { get; set; }
        public string? cc { get; set; }
        public string? date { get; set; }
        public string? filename { get; set; }
        public string? messageId { get; set; }
        public string? recipients { get; set; }
        public string? sender { get; set; }
        public string sentdate { get; set; }
        public DateTime dateTime { get; set; } 
        public string attachments { get; set; }
        public string hasattachments { get; set; }
        public string domain { get; set; }
        public string role { get; set; }

        public string sentByCravens { get; set; }
                

    

    }

    public class TimeEntries
    {
        public HashSet<TimeEntry> entries { get; set;} 

        public TimeEntries()
        {
            entries = new HashSet<TimeEntry>();
        }
    }
}