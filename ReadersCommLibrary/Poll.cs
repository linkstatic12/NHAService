using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReadersCommLibrary
{
    public class Poll
    {
        public Int32 ID { get; set; }
        public DateTime EntryDateTime { get; set; }

        public Poll(Int32 ID, DateTime EntryDateTime)
        {
            this.ID = ID;
            this.EntryDateTime = EntryDateTime;
        }
    }
}
