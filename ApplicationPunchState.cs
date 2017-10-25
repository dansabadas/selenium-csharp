using System;
using System.Runtime.Serialization;

namespace selenium_csharp
{
    [DataContract]
    public class ApplicationPunchState
    {
        [DataMember]
        public DateTime In { get; set; }

        [DataMember]
        public DateTime Out { get; set; }

        public int StartHour { get; set; }

        public int PollingIntervalMinutes { get; set; }

        public int TimeSpentAtOfficeMinutes { get; set; }
    }
}
