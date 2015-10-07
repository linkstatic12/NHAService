using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace TASDownloadService.AttProcessSummary
{
    class SummaryEntity
    {

        public String SummaryDateCriteria = null;
        public DateTime Date = new DateTime();
        public String criterianame = null;
        public int TotalEmps = 0;
        public int PresentEmps = 0;
        public int Absent = 0;
        public int Leave = 0;
        public int Others = 0;
        public int LateIn = 0;
        public int LateOut = 0;
        public int EarlyIn = 0;
        public int EarlyOut = 0;
        public int OverTime = 0;
        public int ShortLeave = 0;
        public int HalfLeave = 0;
        public int OnTime = 0;
        public TimeSpan LateInMins = new TimeSpan();
        public TimeSpan LateOutMins = new TimeSpan();
        public TimeSpan EarlyInMins = new TimeSpan();
        public TimeSpan EarlyOutMins = new TimeSpan();
        public TimeSpan OverTimeMins = new TimeSpan();
        public TimeSpan ExpectedWorkMins = new TimeSpan();
        public TimeSpan ActualWorkMins = new TimeSpan();
        public TimeSpan LossWorkMins = new TimeSpan();
    }
}
