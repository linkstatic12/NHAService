using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace WMSFFService
{
    public static class GlobalSettings
    {
        public static DateTime _dateTime;
        public static TimeSpan ConvertTime(string p)
        {
            try
            {
                string hour = "";
                string min = "";
                int count = 0;
                int chunkSize = 2;
                int stringLength = 4;

                for (int i = 0; i < stringLength; i += chunkSize)
                {
                    count++;
                    if (count == 1)
                    {
                        hour = p.Substring(i, chunkSize);
                    }
                    if (count == 2)
                    {
                        min = p.Substring(i, chunkSize);
                    }
                    if (i + chunkSize > stringLength)
                    {
                        chunkSize = stringLength - i;
                    }
                }
                TimeSpan _currentTime = new TimeSpan(Convert.ToInt32(hour), Convert.ToInt32(min), 00);
                return _currentTime;
            }
            catch (Exception ex)
            {
                return DateTime.Now.TimeOfDay;
            }
        }
    }
}
