using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using WMSFFService;
using TASDownloadService.Model;

namespace TASDownloadService.AttProcessDaily
{
    public class ManualProcess
    {
        #region --Process Daily Attendance--

        TAS2013Entities context = new TAS2013Entities();

        Emp employee = new Emp();

        public void ManualProcessAttendance(DateTime date,List<Emp> emps, List<AttData> _AttDatas)
        {
            BootstrapAttendance(date, emps, _AttDatas);
            DateTime dateEnd = date.AddDays(1);
            List<PollData> unprocessedPolls = context.PollDatas.Where(p => (p.EntDate >= date && p.EntDate <=dateEnd)).OrderBy(e => e.EntTime).ToList();
            foreach (PollData up in unprocessedPolls)
            {
                try
                {
                    //Check AttData with EmpDate
                    if (_AttDatas.Where(attd => attd.EmpDate == up.EmpDate).Count() > 0)
                    {
                        AttData attendanceRecord = _AttDatas.First(attd => attd.EmpDate == up.EmpDate);
                        employee = attendanceRecord.Emp;
                        Shift shift = employee.Shift;
                        //Set Time In and Time Out in AttData
                        if (attendanceRecord.Emp.Shift.OpenShift == true)
                        {
                            //Set Time In and Time Out for open shift
                            CalculateTimeINOUTOpenShift(attendanceRecord, up);
                        }
                        else
                        {
                            TimeSpan checkTimeEnd = new TimeSpan();
                            DateTime TimeInCheck = new DateTime();
                            if (attendanceRecord.TimeIn == null)
                            {
                                TimeInCheck = attendanceRecord.AttDate.Value.Add(attendanceRecord.DutyTime.Value);
                            }
                            else
                                TimeInCheck = attendanceRecord.TimeIn.Value;
                            if (attendanceRecord.ShifMin == 0)
                                checkTimeEnd = TimeInCheck.TimeOfDay.Add(new TimeSpan(0, 480, 0));
                            else
                                checkTimeEnd = TimeInCheck.TimeOfDay.Add(new TimeSpan(0, (int)attendanceRecord.ShifMin, 0));
                            if (checkTimeEnd.Days > 0)
                            {
                                //if Time out occur at next day
                                if (up.RdrDuty == 5)
                                {
                                    DateTime dt = new DateTime();
                                    dt = up.EntDate.Date.AddDays(-1);
                                    var _attData = context.AttDatas.FirstOrDefault(aa => aa.AttDate == dt && aa.EmpID == up.EmpID);
                                    if (_attData != null)
                                    {

                                        if (_attData.TimeIn != null)
                                        {
                                            if ((up.EntTime - _attData.TimeIn.Value).Hours < 18)
                                            {
                                                attendanceRecord = _attData;
                                                up.EmpDate = up.EmpID.ToString() + dt.Date.ToString("yyMMdd");
                                            }

                                        }
                                        else
                                        {
                                            attendanceRecord = _attData;
                                            up.EmpDate = up.EmpID.ToString() + dt.Date.ToString("yyMMdd");
                                        }

                                    }
                                }
                            }
                            //Set Time In and Time Out
                            //Set Time In and Time Out
                            if (up.RdrDuty == 5)
                            {
                                if (attendanceRecord.TimeIn != null)
                                {
                                    TimeSpan dt = (TimeSpan)(up.EntTime.TimeOfDay - attendanceRecord.TimeIn.Value.TimeOfDay);
                                    if (dt.Minutes < 0)
                                    {
                                        DateTime dt1 = new DateTime();
                                        dt1 = up.EntDate.Date.AddDays(-1);
                                        var _attData = context.AttDatas.FirstOrDefault(aa => aa.AttDate == dt1 && aa.EmpID == up.EmpID);
                                        attendanceRecord = _attData;
                                        up.EmpDate = up.EmpID.ToString() + dt1.Date.ToString("yyMMdd");
                                        CalculateTimeINOUT(attendanceRecord, up);
                                    }
                                    else
                                        CalculateTimeINOUT(attendanceRecord, up);
                                }
                                else
                                    CalculateTimeINOUT(attendanceRecord, up);
                            }
                            else
                                CalculateTimeINOUT(attendanceRecord, up);
                        }
                        if (employee.Shift.OpenShift == true)
                        {
                            if (up.EntTime.TimeOfDay < OpenShiftThresholdEnd)
                            {
                                DateTime dt = up.EntDate.Date.AddDays(-1);
                                CalculateOpenShiftTimes(context.AttDatas.FirstOrDefault(aa => aa.AttDate == dt && aa.EmpID == up.EmpID), shift);
                            }
                        }
                        //If TimeIn and TimeOut are not null, then calculate other Atributes
                        if (context.AttDatas.First(attd => attd.EmpDate == up.EmpDate).TimeIn != null && context.AttDatas.First(attd => attd.EmpDate == up.EmpDate).TimeOut != null)
                        {
                            if (context.Rosters.Where(r => r.EmpDate == up.EmpDate).Count() > 0)
                            {
                                CalculateRosterTimes(attendanceRecord, context.Rosters.FirstOrDefault(r => r.EmpDate == up.EmpDate), shift);
                            }
                            else
                            {
                                if (shift.OpenShift == true)
                                {
                                    if (up.EntTime.TimeOfDay < OpenShiftThresholdEnd)
                                    {
                                        DateTime dt = up.EntDate.Date.AddDays(-1);
                                        CalculateOpenShiftTimes(context.AttDatas.FirstOrDefault(aa => aa.AttDate == dt && aa.EmpID == up.EmpID), shift);
                                        CalculateOpenShiftTimes(attendanceRecord, shift);
                                    }
                                    else
                                    {
                                        //Calculate open shifft time of the same date
                                        CalculateOpenShiftTimes(attendanceRecord, shift);
                                    }
                                }
                                else
                                {
                                    CalculateShiftTimes(attendanceRecord, shift);
                                }
                            }
                        }
                        up.Process = true;
                    }
                }
                catch (Exception ex)
                {
                    string _error = "";
                    if (ex.InnerException.Message != null)
                        _error = ex.InnerException.Message;
                    else
                        _error = ex.Message;
                    _myHelperClass.WriteToLogFile("Attendance Processing Error Level 1 " + _error);
                }
                context.SaveChanges();
            }
            _myHelperClass.WriteToLogFile("Attendance Processing Completed");
            context.Dispose();
        }

        MyCustomFunctions _myHelperClass = new MyCustomFunctions();

        TimeSpan OpenShiftThresholdStart = new TimeSpan(17, 00, 00);
        
        TimeSpan OpenShiftThresholdEnd = new TimeSpan(11, 00, 00);

        #region -- Calculate Time In/Out --
        private void CalculateTimeINOUTOpenShift(AttData attendanceRecord, PollData up)
        {
            try
            {
                switch (up.RdrDuty)
                {
                    case 1: //IN
                        if (attendanceRecord.Tin0 == null)
                        {
                            if (up.EntTime.TimeOfDay < OpenShiftThresholdEnd)
                            {
                                DateTime dt = new DateTime();
                                dt = up.EntDate.Date.AddDays(-1);
                                var _attData = context.AttDatas.FirstOrDefault(aa => aa.AttDate == dt && aa.EmpID == up.EmpID);
                                if (_attData != null)
                                {

                                    if (_attData.TimeIn != null)
                                    {
                                        if (_attData.TimeIn.Value.TimeOfDay > OpenShiftThresholdStart)
                                        {
                                            //attdata - 1 . multipleTimeIn =  up.EntTime 

                                        }
                                        else
                                        {
                                            attendanceRecord.Tin0 = up.EntTime;
                                            attendanceRecord.TimeIn = up.EntTime;
                                            attendanceRecord.StatusAB = false;
                                            attendanceRecord.StatusP = true;
                                            attendanceRecord.Remarks = null;
                                            attendanceRecord.StatusIN = true;
                                        }
                                    }
                                    else
                                    {
                                        attendanceRecord.Tin0 = up.EntTime;
                                        attendanceRecord.TimeIn = up.EntTime;
                                        attendanceRecord.StatusAB = false;
                                        attendanceRecord.StatusP = true;
                                        attendanceRecord.Remarks = null;
                                        attendanceRecord.StatusIN = true;
                                    }
                                }
                                else
                                {
                                    attendanceRecord.Tin0 = up.EntTime;
                                    attendanceRecord.TimeIn = up.EntTime;
                                    attendanceRecord.StatusAB = false;
                                    attendanceRecord.StatusP = true;
                                    attendanceRecord.Remarks = null;
                                    attendanceRecord.StatusIN = true;
                                }
                            }
                            else
                            {
                                attendanceRecord.Tin0 = up.EntTime;
                                attendanceRecord.TimeIn = up.EntTime;
                                attendanceRecord.StatusAB = false;
                                attendanceRecord.StatusP = true;
                                attendanceRecord.Remarks = null;
                                attendanceRecord.StatusIN = true;
                            }

                        }
                        else if (attendanceRecord.Tin1 == null)
                        {
                            attendanceRecord.Tin1 = up.EntTime;
                            SortingInTime(attendanceRecord);
                        }
                        else if (attendanceRecord.Tin2 == null)
                        {
                            attendanceRecord.Tin2 = up.EntTime;
                            SortingInTime(attendanceRecord);
                        }
                        else if (attendanceRecord.Tin3 == null)
                        {
                            attendanceRecord.Tin3 = up.EntTime;
                            SortingInTime(attendanceRecord);
                        }
                        else if (attendanceRecord.Tin4 == null)
                        {
                            attendanceRecord.Tin4 = up.EntTime;
                            SortingInTime(attendanceRecord);
                        }
                        else if (attendanceRecord.Tin5 == null)
                        {
                            attendanceRecord.Tin5 = up.EntTime;
                            SortingInTime(attendanceRecord);
                        }
                        else if (attendanceRecord.Tin6 == null)
                        {
                            attendanceRecord.Tin6 = up.EntTime;
                            SortingInTime(attendanceRecord);
                        }
                        else if (attendanceRecord.Tin7 == null)
                        {
                            attendanceRecord.Tin7 = up.EntTime;
                            SortingInTime(attendanceRecord);
                        }
                        else if (attendanceRecord.Tin8 == null)
                        {
                            attendanceRecord.Tin8 = up.EntTime;
                            SortingInTime(attendanceRecord);
                        }
                        else if (attendanceRecord.Tin9 == null)
                        {
                            attendanceRecord.Tin9 = up.EntTime;
                            SortingInTime(attendanceRecord);
                        }
                        else if (attendanceRecord.Tin10 == null)
                        {
                            attendanceRecord.Tin10 = up.EntTime;
                            SortingInTime(attendanceRecord);
                        }
                        else
                        {
                            attendanceRecord.Tin11 = up.EntTime;
                            SortingInTime(attendanceRecord);
                        }
                        break;
                    case 5: //OUT
                        if (up.EntTime.TimeOfDay < OpenShiftThresholdEnd)
                        {
                            DateTime dt = up.EntDate.AddDays(-1);
                            if (context.AttDatas.Where(aa => aa.AttDate == dt && aa.EmpID == up.EmpID).Count() > 0)
                            {
                                AttData AttDataOfPreviousDay = context.AttDatas.FirstOrDefault(aa => aa.AttDate == dt && aa.EmpID == up.EmpID);
                                if (AttDataOfPreviousDay.TimeIn != null)
                                {
                                    if (AttDataOfPreviousDay.TimeIn.Value.TimeOfDay > OpenShiftThresholdStart)
                                    {
                                        //AttDate -1, Possible TimeOut = up.entryTime
                                        MarkOUTForOpenShift(up.EntTime, AttDataOfPreviousDay);
                                    }
                                    else
                                    {
                                        // Mark as out of that day
                                        MarkOUTForOpenShift(up.EntTime, attendanceRecord);
                                    }
                                }
                                else
                                    MarkOUTForOpenShift(up.EntTime, attendanceRecord);
                            }
                            else
                            {
                                // Mark as out of that day
                                MarkOUTForOpenShift(up.EntTime, attendanceRecord);
                            }


                        }
                        else
                        {
                            //Mark as out of that day
                            MarkOUTForOpenShift(up.EntTime, attendanceRecord);
                        }
                        //-------------------------------------------------------
                        context.SaveChanges();
                        break;
                }
            }
            catch (Exception ex)
            {
                string _error = "";
                if (ex.InnerException.Message != null)
                    _error = ex.InnerException.Message;
                else
                    _error = ex.Message;
                _myHelperClass.WriteToLogFile("Attendance Processing Error at Markin In/Out " + _error);
            }
        }

        private void MarkOUTForOpenShift(DateTime _pollTime, AttData _attendanceRecord)
        {
            if (_attendanceRecord.Tout0 == null)
            {
                _attendanceRecord.Tout0 = _pollTime;
                _attendanceRecord.TimeOut = _pollTime;
                SortingOutTime(_attendanceRecord);
            }
            else if (_attendanceRecord.Tout1 == null)
            {
                _attendanceRecord.Tout1 = _pollTime;
                _attendanceRecord.TimeOut = _pollTime;
                SortingOutTime(_attendanceRecord);
            }
            else if (_attendanceRecord.Tout2 == null)
            {
                _attendanceRecord.Tout2 = _pollTime;
                _attendanceRecord.TimeOut = _pollTime;
                SortingOutTime(_attendanceRecord);
            }
            else if (_attendanceRecord.Tout3 == null)
            {
                _attendanceRecord.Tout3 = _pollTime;
                _attendanceRecord.TimeOut = _pollTime;
                SortingOutTime(_attendanceRecord);
            }
            else
            {
                _attendanceRecord.Tout4 = _pollTime;
                _attendanceRecord.TimeOut = _pollTime;
                SortingOutTime(_attendanceRecord);
            }
            //else if (_attendanceRecord.Tout5 == null)
            //{
            //    _attendanceRecord.Tout5 = up.EntTime;
            //    _attendanceRecord.TimeOut = up.EntTime;
            //    SortingOutTime(_attendanceRecord);
            //}
            //else if (_attendanceRecord.Tout6 == null)
            //{
            //    _attendanceRecord.Tout6 = up.EntTime;
            //    _attendanceRecord.TimeOut = up.EntTime;
            //    SortingOutTime(_attendanceRecord);
            //}
            //else if (_attendanceRecord.Tout7 == null)
            //{
            //    _attendanceRecord.Tout7 = up.EntTime;
            //    _attendanceRecord.TimeOut = up.EntTime;
            //    SortingOutTime(_attendanceRecord);
            //}
            //else if (_attendanceRecord.Tout8 == null)
            //{
            //    _attendanceRecord.Tout8 = up.EntTime;
            //    _attendanceRecord.TimeOut = up.EntTime;
            //    SortingOutTime(_attendanceRecord);
            //}
            //else if (_attendanceRecord.Tout9 == null)
            //{
            //    _attendanceRecord.Tout9 = up.EntTime;
            //    _attendanceRecord.TimeOut = up.EntTime;
            //    SortingOutTime(_attendanceRecord);
            //}
            //else if (_attendanceRecord.Tout10 == null)
            //{
            //    _attendanceRecord.Tout10 = up.EntTime;
            //    _attendanceRecord.TimeOut = up.EntTime;
            //    SortingOutTime(_attendanceRecord);
            //}
            //else
            //{
            //    _attendanceRecord.Tout11 = up.EntTime;
            //    _attendanceRecord.TimeOut = up.EntTime;
            //    SortingOutTime(_attendanceRecord);
            //}
        }

        private void CalculateTimeINOUT(AttData attendanceRecord, PollData up)
        {
            try
            {
                switch (up.RdrDuty)
                {
                    case 1: //IN
                        if (attendanceRecord.Tin0 == null)
                        {
                            attendanceRecord.Tin0 = up.EntTime;
                            attendanceRecord.TimeIn = up.EntTime;
                            attendanceRecord.StatusAB = false;
                            attendanceRecord.StatusP = true;
                            attendanceRecord.Remarks = null;
                        }
                        else if (attendanceRecord.Tin1 == null)
                        {
                            attendanceRecord.Tin1 = up.EntTime;
                            SortingInTime(attendanceRecord);
                        }
                        else if (attendanceRecord.Tin2 == null)
                        {
                            attendanceRecord.Tin2 = up.EntTime;
                            SortingInTime(attendanceRecord);
                        }
                        else if (attendanceRecord.Tin3 == null)
                        {
                            attendanceRecord.Tin3 = up.EntTime;
                            SortingInTime(attendanceRecord);
                        }
                        else if (attendanceRecord.Tin4 == null)
                        {
                            attendanceRecord.Tin4 = up.EntTime;
                            SortingInTime(attendanceRecord);
                        }
                        else if (attendanceRecord.Tin5 == null)
                        {
                            attendanceRecord.Tin5 = up.EntTime;
                            SortingInTime(attendanceRecord);
                        }
                        else if (attendanceRecord.Tin6 == null)
                        {
                            attendanceRecord.Tin6 = up.EntTime;
                            SortingInTime(attendanceRecord);
                        }
                        else if (attendanceRecord.Tin7 == null)
                        {
                            attendanceRecord.Tin7 = up.EntTime;
                            SortingInTime(attendanceRecord);
                        }
                        else if (attendanceRecord.Tin8 == null)
                        {
                            attendanceRecord.Tin8 = up.EntTime;
                            SortingInTime(attendanceRecord);
                        }
                        else if (attendanceRecord.Tin9 == null)
                        {
                            attendanceRecord.Tin9 = up.EntTime;
                            SortingInTime(attendanceRecord);
                        }
                        else if (attendanceRecord.Tin10 == null)
                        {
                            attendanceRecord.Tin10 = up.EntTime;
                            SortingInTime(attendanceRecord);
                        }
                        else
                        {
                            attendanceRecord.Tin11 = up.EntTime;
                            SortingInTime(attendanceRecord);
                        }
                        break;
                    case 5: //OUT
                        if (attendanceRecord.Tout0 == null)
                        {
                            attendanceRecord.Tout0 = up.EntTime;
                            attendanceRecord.TimeOut = up.EntTime;
                            SortingOutTime(attendanceRecord);
                        }
                        else if (attendanceRecord.Tout1 == null)
                        {
                            attendanceRecord.Tout1 = up.EntTime;
                            attendanceRecord.TimeOut = up.EntTime;
                            SortingOutTime(attendanceRecord);
                        }
                        else if (attendanceRecord.Tout2 == null)
                        {
                            attendanceRecord.Tout2 = up.EntTime;
                            attendanceRecord.TimeOut = up.EntTime;
                            SortingOutTime(attendanceRecord);
                        }
                        else if (attendanceRecord.Tout3 == null)
                        {
                            attendanceRecord.Tout3 = up.EntTime;
                            attendanceRecord.TimeOut = up.EntTime;
                            SortingOutTime(attendanceRecord);
                        }
                        else if (attendanceRecord.Tout4 == null)
                        {
                            attendanceRecord.Tout4 = up.EntTime;
                            attendanceRecord.TimeOut = up.EntTime;
                            SortingOutTime(attendanceRecord);
                        }
                        else if (attendanceRecord.Tout5 == null)
                        {
                            attendanceRecord.Tout5 = up.EntTime;
                            attendanceRecord.TimeOut = up.EntTime;
                            SortingOutTime(attendanceRecord);
                        }
                        else if (attendanceRecord.Tout6 == null)
                        {
                            attendanceRecord.Tout6 = up.EntTime;
                            attendanceRecord.TimeOut = up.EntTime;
                            SortingOutTime(attendanceRecord);
                        }
                        else if (attendanceRecord.Tout7 == null)
                        {
                            attendanceRecord.Tout7 = up.EntTime;
                            attendanceRecord.TimeOut = up.EntTime;
                            SortingOutTime(attendanceRecord);
                        }
                        else if (attendanceRecord.Tout8 == null)
                        {
                            attendanceRecord.Tout8 = up.EntTime;
                            attendanceRecord.TimeOut = up.EntTime;
                            SortingOutTime(attendanceRecord);
                        }
                        else if (attendanceRecord.Tout9 == null)
                        {
                            attendanceRecord.Tout9 = up.EntTime;
                            attendanceRecord.TimeOut = up.EntTime;
                            SortingOutTime(attendanceRecord);
                        }
                        else if (attendanceRecord.Tout10 == null)
                        {
                            attendanceRecord.Tout10 = up.EntTime;
                            attendanceRecord.TimeOut = up.EntTime;
                            SortingOutTime(attendanceRecord);
                        }
                        else
                        {
                            attendanceRecord.Tout11 = up.EntTime;
                            attendanceRecord.TimeOut = up.EntTime;
                            SortingOutTime(attendanceRecord);
                        }
                        break;
                    case 8: //DUTY
                        if (attendanceRecord.Tin0 != null)
                        {
                            if (attendanceRecord.Tout0 == null)
                            {
                                attendanceRecord.Tout0 = up.EntTime;
                                attendanceRecord.TimeOut = up.EntTime;
                            }
                            else if (attendanceRecord.Tin1 == null)
                            {
                                attendanceRecord.Tin1 = up.EntTime;
                                attendanceRecord.TimeOut = up.EntTime;
                            }
                            else if (attendanceRecord.Tout1 == null)
                            {
                                attendanceRecord.Tout1 = up.EntTime;
                                attendanceRecord.TimeOut = up.EntTime;
                            }
                            else if (attendanceRecord.Tin2 == null)
                            {
                                attendanceRecord.Tin2 = up.EntTime;
                                attendanceRecord.TimeOut = up.EntTime;
                            }
                            else if (attendanceRecord.Tout2 == null)
                            {
                                attendanceRecord.Tout2 = up.EntTime;
                                attendanceRecord.TimeOut = up.EntTime;
                            }
                            else if (attendanceRecord.Tin3 == null)
                            {
                                attendanceRecord.Tin3 = up.EntTime;
                                attendanceRecord.TimeOut = up.EntTime;
                            }
                            else if (attendanceRecord.Tout3 == null)
                            {
                                attendanceRecord.Tout3 = up.EntTime;
                                attendanceRecord.TimeOut = up.EntTime;
                            }
                            else if (attendanceRecord.Tin4 == null)
                            {
                                attendanceRecord.Tin4 = up.EntTime;
                                attendanceRecord.TimeOut = up.EntTime;
                            }
                            else if (attendanceRecord.Tout4 == null)
                            {
                                attendanceRecord.Tout4 = up.EntTime;
                                attendanceRecord.TimeOut = up.EntTime;
                            }
                            else if (attendanceRecord.Tin5 == null)
                            {
                                attendanceRecord.Tin5 = up.EntTime;
                                attendanceRecord.TimeOut = up.EntTime;
                            }
                            else if (attendanceRecord.Tout5 == null)
                            {
                                attendanceRecord.Tout5 = up.EntTime;
                                attendanceRecord.TimeOut = up.EntTime;
                            }
                            else if (attendanceRecord.Tin6 == null)
                            {
                                attendanceRecord.Tin6 = up.EntTime;
                                attendanceRecord.TimeOut = up.EntTime;
                            }
                            else if (attendanceRecord.Tout6 == null)
                            {
                                attendanceRecord.Tout6 = up.EntTime;
                                attendanceRecord.TimeOut = up.EntTime;
                            }
                            //
                            else if (attendanceRecord.Tin7 == null)
                            {
                                attendanceRecord.Tin7 = up.EntTime;
                                attendanceRecord.TimeOut = up.EntTime;
                            }
                            else if (attendanceRecord.Tout7 == null)
                            {
                                attendanceRecord.Tout7 = up.EntTime;
                                attendanceRecord.TimeOut = up.EntTime;
                            }
                            else if (attendanceRecord.Tin8 == null)
                            {
                                attendanceRecord.Tin8 = up.EntTime;
                                attendanceRecord.TimeOut = up.EntTime;
                            }
                            else if (attendanceRecord.Tout8 == null)
                            {
                                attendanceRecord.Tout8 = up.EntTime;
                                attendanceRecord.TimeOut = up.EntTime;
                            }
                            else if (attendanceRecord.Tin9 == null)
                            {
                                attendanceRecord.Tin9 = up.EntTime;
                                attendanceRecord.TimeOut = up.EntTime;
                            }
                            else if (attendanceRecord.Tout9 == null)
                            {
                                attendanceRecord.Tout9 = up.EntTime;
                                attendanceRecord.TimeOut = up.EntTime;
                            }
                            else if (attendanceRecord.Tin10 == null)
                            {
                                attendanceRecord.Tin10 = up.EntTime;
                                attendanceRecord.TimeOut = up.EntTime;
                            }
                            else if (attendanceRecord.Tout10 == null)
                            {
                                attendanceRecord.Tout10 = up.EntTime;
                                attendanceRecord.TimeOut = up.EntTime;
                            }
                            else if (attendanceRecord.Tin11 == null)
                            {
                                attendanceRecord.Tin11 = up.EntTime;
                                attendanceRecord.TimeOut = up.EntTime;
                            }
                            else if (attendanceRecord.Tout11 == null)
                            {
                                attendanceRecord.Tout11 = up.EntTime;
                                attendanceRecord.TimeOut = up.EntTime;
                            }
                            else if (attendanceRecord.Tin12 == null)
                            {
                                attendanceRecord.Tin12 = up.EntTime;
                                attendanceRecord.TimeOut = up.EntTime;
                            }
                            else if (attendanceRecord.Tout12 == null)
                            {
                                attendanceRecord.Tout12 = up.EntTime;
                                attendanceRecord.TimeOut = up.EntTime;
                            }
                            else if (attendanceRecord.Tin13 == null)
                            {
                                attendanceRecord.Tin13 = up.EntTime;
                                attendanceRecord.TimeOut = up.EntTime;
                            }
                            else
                            {
                                attendanceRecord.Tout13 = up.EntTime;
                                attendanceRecord.TimeOut = up.EntTime;
                            }
                        }
                        else
                        {
                            attendanceRecord.Tin0 = up.EntTime;
                            attendanceRecord.TimeIn = up.EntTime;
                            attendanceRecord.StatusAB = false;
                            attendanceRecord.StatusP = true;
                            attendanceRecord.Remarks = null;
                        }
                        break;
                }
            }
            catch (Exception ex)
            {
                //Error in TimeIN/OUT
                _myHelperClass.WriteToLogFile("Error At Creating Attendance");
            }
        }

        // Sorting Time In
        private void SortingInTime(AttData attendanceRecord)
        {
            List<DateTime> _InTimes = new List<DateTime>();

            if (attendanceRecord.Tin0 != null)
                _InTimes.Add((DateTime)attendanceRecord.Tin0);
            if (attendanceRecord.Tin1 != null)
                _InTimes.Add((DateTime)attendanceRecord.Tin1);
            if (attendanceRecord.Tin2 != null)
                _InTimes.Add((DateTime)attendanceRecord.Tin2);
            if (attendanceRecord.Tin3 != null)
                _InTimes.Add((DateTime)attendanceRecord.Tin3);
            if (attendanceRecord.Tin4 != null)
                _InTimes.Add((DateTime)attendanceRecord.Tin4);
            if (attendanceRecord.Tin5 != null)
                _InTimes.Add((DateTime)attendanceRecord.Tin5);
            if (attendanceRecord.Tin6 != null)
                _InTimes.Add((DateTime)attendanceRecord.Tin6);
            if (attendanceRecord.Tin7 != null)
                _InTimes.Add((DateTime)attendanceRecord.Tin7);
            if (attendanceRecord.Tin8 != null)
                _InTimes.Add((DateTime)attendanceRecord.Tin8);
            if (attendanceRecord.Tin9 != null)
                _InTimes.Add((DateTime)attendanceRecord.Tin9);
            if (attendanceRecord.Tin10 != null)
                _InTimes.Add((DateTime)attendanceRecord.Tin10);

            var list = _InTimes.OrderBy(x => x.TimeOfDay.Hours).ToList();
            PlacedSortedInTime(attendanceRecord, list);

        }

        private void PlacedSortedInTime(AttData attendanceRecord, List<DateTime> _InTimes)
        {
            for (int i = 0; i < _InTimes.Count; i++)
            {
                switch (i)
                {
                    case 0:
                        attendanceRecord.Tin0 = _InTimes[i];
                        attendanceRecord.TimeIn = _InTimes[i];
                        break;
                    case 1:
                        attendanceRecord.Tin1 = _InTimes[i];
                        break;
                    case 2:
                        attendanceRecord.Tin2 = _InTimes[i];
                        break;
                    case 3:
                        attendanceRecord.Tin3 = _InTimes[i];
                        break;
                    case 4:
                        attendanceRecord.Tin4 = _InTimes[i];
                        break;
                    case 5:
                        attendanceRecord.Tin5 = _InTimes[i];
                        break;
                    case 6:
                        attendanceRecord.Tin6 = _InTimes[i];
                        break;
                    case 7:
                        attendanceRecord.Tin7 = _InTimes[i];
                        break;
                    case 8:
                        attendanceRecord.Tin8 = _InTimes[i];
                        break;
                    case 9:
                        attendanceRecord.Tin9 = _InTimes[i];
                        break;
                }
            }
        }
        //Sorting Time Out
        private void SortingOutTime(AttData attendanceRecord)
        {
            List<DateTime> _OutTimes = new List<DateTime>();

            if (attendanceRecord.Tout0 != null)
                _OutTimes.Add((DateTime)attendanceRecord.Tout0);
            if (attendanceRecord.Tout1 != null)
                _OutTimes.Add((DateTime)attendanceRecord.Tout1);
            if (attendanceRecord.Tout2 != null)
                _OutTimes.Add((DateTime)attendanceRecord.Tout2);
            if (attendanceRecord.Tout3 != null)
                _OutTimes.Add((DateTime)attendanceRecord.Tout3);
            if (attendanceRecord.Tout4 != null)
                _OutTimes.Add((DateTime)attendanceRecord.Tout4);
            if (attendanceRecord.Tout5 != null)
                _OutTimes.Add((DateTime)attendanceRecord.Tout5);
            if (attendanceRecord.Tout6 != null)
                _OutTimes.Add((DateTime)attendanceRecord.Tout6);
            if (attendanceRecord.Tout7 != null)
                _OutTimes.Add((DateTime)attendanceRecord.Tout7);
            if (attendanceRecord.Tout8 != null)
                _OutTimes.Add((DateTime)attendanceRecord.Tout8);
            if (attendanceRecord.Tout9 != null)
                _OutTimes.Add((DateTime)attendanceRecord.Tout9);
            if (attendanceRecord.Tout10 != null)
                _OutTimes.Add((DateTime)attendanceRecord.Tout10);

            var list = _OutTimes.OrderBy(x => x.TimeOfDay.Hours).ToList();
            PlacedSortedOutTime(attendanceRecord, list);


        }

        private void PlacedSortedOutTime(AttData attendanceRecord, List<DateTime> _OutTimes)
        {
            for (int i = 0; i < _OutTimes.Count; i++)
            {
                switch (i)
                {
                    case 0:
                        attendanceRecord.Tout0 = _OutTimes[i];
                        attendanceRecord.TimeOut = _OutTimes[i];
                        break;
                    case 1:
                        attendanceRecord.Tout1 = _OutTimes[i];
                        attendanceRecord.TimeOut = _OutTimes[i];
                        break;
                    case 2:
                        attendanceRecord.Tout2 = _OutTimes[i];
                        attendanceRecord.TimeOut = _OutTimes[i];
                        break;
                    case 3:
                        attendanceRecord.Tout3 = _OutTimes[i];
                        attendanceRecord.TimeOut = _OutTimes[i];
                        break;
                    case 4:
                        attendanceRecord.Tout4 = _OutTimes[i];
                        attendanceRecord.TimeOut = _OutTimes[i];
                        break;
                    case 5:
                        attendanceRecord.Tout5 = _OutTimes[i];
                        attendanceRecord.TimeOut = _OutTimes[i];
                        break;
                    case 6:
                        attendanceRecord.Tout6 = _OutTimes[i];
                        attendanceRecord.TimeOut = _OutTimes[i];
                        break;
                    case 7:
                        attendanceRecord.Tout7 = _OutTimes[i];
                        attendanceRecord.TimeOut = _OutTimes[i];
                        break;
                    case 8:
                        attendanceRecord.Tout8 = _OutTimes[i];
                        attendanceRecord.TimeOut = _OutTimes[i];
                        break;
                    case 9:
                        attendanceRecord.Tout9 = _OutTimes[i];
                        attendanceRecord.TimeOut = _OutTimes[i];
                        break;
                }
            }
        }
        #endregion

        #region -- Calculate Work Times --

        private void CalculateShiftTimes(AttData attendanceRecord, Shift shift)
        {
            try
            {
                //Calculate WorkMin
                attendanceRecord.Remarks = "";
                TimeSpan mins = (TimeSpan)(attendanceRecord.TimeOut - attendanceRecord.TimeIn);
                //Check if GZ holiday then place all WorkMin in GZOTMin
                if (attendanceRecord.StatusGZ == true)
                {
                    attendanceRecord.GZOTMin = (short)mins.TotalMinutes;
                    attendanceRecord.WorkMin = (short)mins.TotalMinutes;
                    attendanceRecord.StatusGZOT = true;
                    attendanceRecord.Remarks = attendanceRecord.Remarks + "[G-OT]";
                }
                //if Rest day then place all WorkMin in OTMin
                else if (attendanceRecord.StatusDO == true)
                {
                    attendanceRecord.OTMin = (short)mins.TotalMinutes;
                    attendanceRecord.WorkMin = (short)mins.TotalMinutes;
                    attendanceRecord.StatusOT = true;
                    attendanceRecord.Remarks = attendanceRecord.Remarks + "[R-OT]";

                }
                else
                {
                    /////////// to-do -----calculate Margins for those shifts which has break mins 
                    if (shift.HasBreak == true)
                    {
                        attendanceRecord.WorkMin = (short)(mins.TotalMinutes - shift.BreakMin);
                        attendanceRecord.ShifMin = (short)(CalculateShiftMinutes(shift, attendanceRecord.AttDate.Value.DayOfWeek) - (short)shift.BreakMin);
                    }
                    else
                    {
                        //Calculate Late IN, Compare margin with Shift Late In
                        if (attendanceRecord.TimeIn.Value.TimeOfDay > attendanceRecord.DutyTime)
                        {
                            TimeSpan lateMinsSpan = (TimeSpan)(attendanceRecord.TimeIn.Value.TimeOfDay - attendanceRecord.DutyTime);
                            if (lateMinsSpan.TotalMinutes > shift.LateIn)
                            {
                                attendanceRecord.LateIn = (short)lateMinsSpan.TotalMinutes;
                                attendanceRecord.StatusLI = true;
                                attendanceRecord.EarlyIn = null;
                                attendanceRecord.Remarks = attendanceRecord.Remarks + "[LI]";
                            }
                            else
                            {
                                attendanceRecord.StatusLI = null;
                                attendanceRecord.LateIn = null;
                                attendanceRecord.Remarks.Replace("[LI]", "");
                            }
                        }
                        else
                        {
                            attendanceRecord.StatusLI = null;
                            attendanceRecord.LateIn = null;
                            attendanceRecord.Remarks.Replace("[LI]", "");
                        }

                        //Calculate Early In, Compare margin with Shift Early In
                        if (attendanceRecord.TimeIn.Value.TimeOfDay < attendanceRecord.DutyTime)
                        {
                            TimeSpan EarlyInMinsSpan = (TimeSpan)(attendanceRecord.DutyTime - attendanceRecord.TimeIn.Value.TimeOfDay);
                            if (EarlyInMinsSpan.TotalMinutes > shift.EarlyIn)
                            {
                                attendanceRecord.EarlyIn = (short)EarlyInMinsSpan.TotalMinutes;
                                attendanceRecord.StatusEI = true;
                                attendanceRecord.LateIn = null;
                                attendanceRecord.Remarks = attendanceRecord.Remarks + "[EI]";
                            }
                            else
                            {
                                attendanceRecord.StatusEI = null;
                                attendanceRecord.EarlyIn = null;
                                attendanceRecord.Remarks.Replace("[EI]", "");
                            }
                        }
                        else
                        {
                            attendanceRecord.StatusEI = null;
                            attendanceRecord.EarlyIn = null;
                            attendanceRecord.Remarks.Replace("[EI]", "");
                        }

                        // CalculateShiftEndTime = ShiftStart + DutyHours
                        DateTime shiftEnd = CalculateShiftEndTime(shift, attendanceRecord.AttDate.Value, attendanceRecord.DutyTime.Value);

                        //Calculate Early Out, Compare margin with Shift Early Out
                        if (attendanceRecord.TimeOut < shiftEnd)
                        {
                            TimeSpan EarlyOutMinsSpan = (TimeSpan)(shiftEnd - attendanceRecord.TimeOut);
                            if (EarlyOutMinsSpan.TotalMinutes > shift.EarlyOut)
                            {
                                attendanceRecord.EarlyOut = (short)EarlyOutMinsSpan.TotalMinutes;
                                attendanceRecord.StatusEO = true;
                                attendanceRecord.LateOut = null;
                                attendanceRecord.Remarks = attendanceRecord.Remarks + "[EO]";
                            }
                            else
                            {
                                attendanceRecord.StatusEO = null;
                                attendanceRecord.EarlyOut = null;
                                attendanceRecord.Remarks.Replace("[EO]", "");
                            }
                        }
                        else
                        {
                            attendanceRecord.StatusEO = null;
                            attendanceRecord.EarlyOut = null;
                            attendanceRecord.Remarks.Replace("[EO]", "");
                        }
                        //Calculate Late Out, Compare margin with Shift Late Out
                        if (attendanceRecord.TimeOut > shiftEnd)
                        {
                            TimeSpan LateOutMinsSpan = (TimeSpan)(attendanceRecord.TimeOut - shiftEnd);
                            if (LateOutMinsSpan.TotalMinutes > shift.LateOut)
                            {
                                attendanceRecord.LateOut = (short)LateOutMinsSpan.TotalMinutes;
                                // Late Out cannot have an early out, In case of poll at multiple times before and after shiftend
                                attendanceRecord.EarlyOut = null;
                                attendanceRecord.StatusLO = true;
                                attendanceRecord.Remarks = attendanceRecord.Remarks + "[LO]";
                            }
                            else
                            {
                                attendanceRecord.StatusLO = null;
                                attendanceRecord.LateOut = null;
                                attendanceRecord.Remarks.Replace("[LO]", "");
                            }
                        }
                        else
                        {
                            attendanceRecord.StatusLO = null;
                            attendanceRecord.LateOut = null;
                            attendanceRecord.Remarks.Replace("[LO]", "");
                        }
                        attendanceRecord.WorkMin = (short)(mins.TotalMinutes);
                        // round off work mins if overtime less than shift.OverTimeMin >
                        //if ((attendanceRecord.WorkMin > CalculateShiftMinutes(shift, attendanceRecord.AttDate.Value.DayOfWeek)) && (attendanceRecord.WorkMin <= (CalculateShiftMinutes(shift, attendanceRecord.AttDate.Value.DayOfWeek) + shift.OverTimeMin)))
                        //{
                        //    attendanceRecord.WorkMin = CalculateShiftMinutes(shift, attendanceRecord.AttDate.Value.DayOfWeek);
                        //}
                        if (attendanceRecord.WorkMin > (attendanceRecord.ShifMin + shift.OverTimeMin))
                        {
                            attendanceRecord.OTMin = (short)(attendanceRecord.WorkMin - attendanceRecord.ShifMin);
                            attendanceRecord.WorkMin = attendanceRecord.ShifMin;
                            attendanceRecord.StatusOT = true;

                        }
                        //if ((attendanceRecord.WorkMin - attendanceRecord.ShifMin) > 0)
                        //{
                        //    attendanceRecord.OTMin = (short)(attendanceRecord.ShifMin - attendanceRecord.WorkMin);
                        //}

                        if ((attendanceRecord.StatusGZ != true || attendanceRecord.StatusDO != true) && employee.HasOT == true)
                        {
                            if (attendanceRecord.LateOut != null)
                            {
                                //attendanceRecord.OTMin = attendanceRecord.LateOut;
                                attendanceRecord.StatusOT = true;
                                attendanceRecord.Remarks = attendanceRecord.Remarks + "[N-OT]";
                            }
                        }
                        //Mark Absent if less than 4 hours
                        if (attendanceRecord.AttDate.Value.DayOfWeek != DayOfWeek.Friday && attendanceRecord.StatusDO != true && attendanceRecord.StatusGZ != true)
                        {
                            short MinShiftMin = (short)shift.MinHrs;
                            if (attendanceRecord.WorkMin < MinShiftMin)
                            {
                                attendanceRecord.StatusAB = true;
                                attendanceRecord.StatusP = false;
                                attendanceRecord.Remarks = "[Absent]";
                            }
                            else
                            {
                                attendanceRecord.StatusAB = false;
                                attendanceRecord.StatusP = true;
                                attendanceRecord.Remarks.Replace("[Absent]", "");
                            }

                        }
                    }
                }
            }
            catch (Exception ex)
            {

            }
        }

        private void CalculateOpenShiftTimes(AttData attendanceRecord, Shift shift)
        {
            try
            {
                //Calculate WorkMin
                if (attendanceRecord != null)
                {
                    if (attendanceRecord.TimeOut != null && attendanceRecord.TimeIn != null)
                    {
                        attendanceRecord.Remarks = "";
                        TimeSpan mins = (TimeSpan)(attendanceRecord.TimeOut - attendanceRecord.TimeIn);
                        //Check if GZ holiday then place all WorkMin in GZOTMin
                        if (attendanceRecord.StatusGZ == true)
                        {
                            attendanceRecord.GZOTMin = (short)mins.TotalMinutes;
                            attendanceRecord.WorkMin = (short)mins.TotalMinutes;
                            attendanceRecord.StatusGZOT = true;
                            attendanceRecord.Remarks = attendanceRecord.Remarks + "[GZ-OT]";
                        }
                        else if (attendanceRecord.StatusDO == true)
                        {
                            attendanceRecord.OTMin = (short)mins.TotalMinutes;
                            attendanceRecord.WorkMin = (short)mins.TotalMinutes;
                            attendanceRecord.StatusOT = true;
                            attendanceRecord.Remarks = attendanceRecord.Remarks + "[R-OT]";
                            // RoundOff Overtime
                            if ((employee.EmpType.CatID == 2 || employee.EmpType.CatID == 4) && employee.CompanyID == 1)
                            {
                                if (attendanceRecord.OTMin > 0)
                                {
                                    float OTmins = (float)attendanceRecord.OTMin;
                                    float remainder = OTmins / 60;
                                    int intpart = (int)remainder;
                                    double fracpart = remainder - intpart;
                                    if (fracpart < 0.5)
                                    {
                                        attendanceRecord.OTMin = (short)(intpart * 60);
                                    }
                                }
                            }
                        }
                        else
                        {
                            if (shift.HasBreak == true)
                            {
                                attendanceRecord.WorkMin = (short)(mins.TotalMinutes - shift.BreakMin);
                                attendanceRecord.ShifMin = (short)(CalculateShiftMinutes(shift, attendanceRecord.AttDate.Value.DayOfWeek) - (short)shift.BreakMin);
                            }
                            else
                            {
                                attendanceRecord.Remarks.Replace("[Absent]", "");
                                attendanceRecord.StatusAB = false;
                                attendanceRecord.StatusP = true;
                                // CalculateShiftEndTime = ShiftStart + DutyHours
                                TimeSpan shiftEnd = CalculateShiftEndTime(shift, attendanceRecord.AttDate.Value.DayOfWeek);
                                attendanceRecord.WorkMin = (short)(mins.TotalMinutes);
                                //Calculate OverTIme, 
                                if ((mins.TotalMinutes > (CalculateShiftMinutes(shift, attendanceRecord.AttDate.Value.DayOfWeek) + shift.OverTimeMin)) && employee.HasOT == true)
                                {
                                    attendanceRecord.OTMin = (Int16)(Convert.ToInt16(mins.TotalMinutes) - CalculateShiftMinutes(shift, attendanceRecord.AttDate.Value.DayOfWeek));
                                    attendanceRecord.WorkMin = (short)((mins.TotalMinutes) - attendanceRecord.OTMin);
                                    attendanceRecord.StatusOT = true;
                                    attendanceRecord.Remarks = attendanceRecord.Remarks + "[N-OT]";
                                }
                                //Calculate Early Out
                                if (mins.TotalMinutes < (CalculateShiftMinutes(shift, attendanceRecord.AttDate.Value.DayOfWeek) - shift.EarlyOut))
                                {
                                    Int16 EarlyoutMin = (Int16)(CalculateShiftMinutes(shift, attendanceRecord.AttDate.Value.DayOfWeek) - Convert.ToInt16(mins.TotalMinutes));
                                    if (EarlyoutMin > shift.EarlyOut)
                                    {
                                        attendanceRecord.EarlyOut = EarlyoutMin;
                                        attendanceRecord.StatusEO = true;
                                        attendanceRecord.Remarks = attendanceRecord.Remarks + "[EO]";
                                    }
                                    else
                                    {
                                        attendanceRecord.StatusEO = null;
                                        attendanceRecord.EarlyOut = null;
                                        attendanceRecord.Remarks.Replace("[EO]", "");
                                    }
                                }
                                else
                                {
                                    attendanceRecord.StatusEO = null;
                                    attendanceRecord.EarlyOut = null;
                                    attendanceRecord.Remarks.Replace("[EO]", "");
                                }
                                // round off work mins if overtime less than shift.OverTimeMin >
                                if (attendanceRecord.WorkMin > CalculateShiftMinutes(shift, attendanceRecord.AttDate.Value.DayOfWeek) && (attendanceRecord.WorkMin <= (CalculateShiftMinutes(shift, attendanceRecord.AttDate.Value.DayOfWeek) + shift.OverTimeMin)))
                                {
                                    attendanceRecord.WorkMin = CalculateShiftMinutes(shift, attendanceRecord.AttDate.Value.DayOfWeek);
                                }
                                // RoundOff Overtime
                                if ((employee.EmpType.CatID == 2 || employee.EmpType.CatID == 4) && employee.CompanyID == 1)
                                {
                                    if (attendanceRecord.OTMin > 0)
                                    {
                                        float OTmins = (float)attendanceRecord.OTMin;
                                        float remainder = OTmins / 60;
                                        int intpart = (int)remainder;
                                        double fracpart = remainder - intpart;
                                        if (fracpart < 0.5)
                                        {
                                            attendanceRecord.OTMin = (short)(intpart * 60);
                                        }
                                    }
                                }
                                //Mark Absent if less than 4 hours
                                if (attendanceRecord.AttDate.Value.DayOfWeek != DayOfWeek.Friday && attendanceRecord.StatusDO != true && attendanceRecord.StatusGZ != true)
                                {
                                    short MinShiftMin = (short)shift.MinHrs;
                                    if (attendanceRecord.WorkMin < MinShiftMin)
                                    {
                                        attendanceRecord.StatusAB = true;
                                        attendanceRecord.StatusP = false;
                                        attendanceRecord.Remarks = attendanceRecord.Remarks + "[Absent]";
                                    }
                                    else
                                    {
                                        attendanceRecord.StatusAB = false;
                                        attendanceRecord.StatusP = true;
                                        attendanceRecord.Remarks.Replace("[Absent]", "");
                                    }

                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                string _error = "";
                if (ex.InnerException.Message != null)
                    _error = ex.InnerException.Message;
                else
                    _error = ex.Message;
                _myHelperClass.WriteToLogFile("Attendance Processing at Calculating Times;  " + _error);
            }
            context.SaveChanges();
        }

        private void CalculateRosterTimes(AttData attendanceRecord, Roster roster, Shift _shift)
        {
            try
            {
                TimeSpan mins = (TimeSpan)(attendanceRecord.TimeOut - attendanceRecord.TimeIn);
                attendanceRecord.Remarks = "";
                if (attendanceRecord.StatusGZ == true)
                {
                    attendanceRecord.GZOTMin = (short)mins.TotalMinutes;
                    attendanceRecord.WorkMin = (short)mins.TotalMinutes;
                    attendanceRecord.StatusGZOT = true;
                    attendanceRecord.Remarks = attendanceRecord.Remarks + "[GZ-OT]";
                }
                else if (attendanceRecord.StatusDO == true)
                {
                    attendanceRecord.OTMin = (short)mins.TotalMinutes;
                    attendanceRecord.WorkMin = (short)mins.TotalMinutes;
                    attendanceRecord.StatusOT = true;
                    attendanceRecord.Remarks = attendanceRecord.Remarks + "[R-OT]";
                    // RoundOff Overtime
                    if ((employee.EmpType.CatID == 2 || employee.EmpType.CatID == 4) && employee.CompanyID == 1)
                    {
                        if (attendanceRecord.OTMin > 0)
                        {
                            float OTmins = (float)attendanceRecord.OTMin;
                            float remainder = OTmins / 60;
                            int intpart = (int)remainder;
                            double fracpart = remainder - intpart;
                            if (fracpart < 0.5)
                            {
                                attendanceRecord.OTMin = (short)(intpart * 60);
                            }
                        }
                    }
                }
                else
                {
                    attendanceRecord.Remarks.Replace("[Absent]", "");
                    attendanceRecord.StatusAB = false;
                    attendanceRecord.StatusP = true;
                    ////------to-do ----------handle shift break time
                    //Calculate Late IN, Compare margin with Shift Late In
                    if (attendanceRecord.TimeIn.Value.TimeOfDay > roster.DutyTime)
                    {
                        TimeSpan lateMinsSpan = (TimeSpan)(attendanceRecord.TimeIn.Value.TimeOfDay - attendanceRecord.DutyTime);
                        if (lateMinsSpan.TotalMinutes > _shift.LateIn)
                        {
                            attendanceRecord.LateIn = (short)lateMinsSpan.TotalMinutes;
                            attendanceRecord.StatusLI = true;
                            attendanceRecord.EarlyIn = null;
                            attendanceRecord.Remarks = attendanceRecord.Remarks + "[LI]";
                        }
                        else
                        {
                            attendanceRecord.LateIn = null;
                            attendanceRecord.StatusLI = null;
                            attendanceRecord.Remarks.Replace("[LI]", "");
                        }
                    }
                    else
                    {
                        attendanceRecord.LateIn = null;
                        attendanceRecord.StatusLI = null;
                        attendanceRecord.Remarks.Replace("[LI]", "");
                    }

                    //Calculate Early In, Compare margin with Shift Early In
                    if (attendanceRecord.TimeIn.Value.TimeOfDay < attendanceRecord.DutyTime)
                    {
                        TimeSpan EarlyInMinsSpan = (TimeSpan)(attendanceRecord.DutyTime - attendanceRecord.TimeIn.Value.TimeOfDay);
                        if (EarlyInMinsSpan.TotalMinutes > _shift.EarlyIn)
                        {
                            attendanceRecord.EarlyIn = (short)EarlyInMinsSpan.TotalMinutes;
                            attendanceRecord.StatusEI = true;
                            attendanceRecord.LateIn = null;
                            attendanceRecord.Remarks = attendanceRecord.Remarks + "[EI]";
                        }
                        else
                        {
                            attendanceRecord.StatusEI = null;
                            attendanceRecord.EarlyIn = null;
                            attendanceRecord.Remarks.Replace("[EI]", "");
                        }
                    }
                    else
                    {
                        attendanceRecord.StatusEI = null;
                        attendanceRecord.EarlyIn = null;
                        attendanceRecord.Remarks.Replace("[EI]", "");
                    }

                    // CalculateShiftEndTime = ShiftStart + DutyHours
                    TimeSpan shiftEnd = (TimeSpan)attendanceRecord.DutyTime + (new TimeSpan(0, (int)roster.WorkMin, 0));

                    //Calculate Early Out, Compare margin with Shift Early Out
                    if (attendanceRecord.TimeOut.Value.TimeOfDay < shiftEnd)
                    {
                        TimeSpan EarlyOutMinsSpan = (TimeSpan)(shiftEnd - attendanceRecord.TimeOut.Value.TimeOfDay);
                        if (EarlyOutMinsSpan.TotalMinutes > _shift.EarlyOut)
                        {
                            attendanceRecord.EarlyOut = (short)EarlyOutMinsSpan.TotalMinutes;
                            attendanceRecord.StatusEO = true;
                            attendanceRecord.LateOut = null;
                            attendanceRecord.Remarks = attendanceRecord.Remarks + "[EO]";
                        }
                        else
                        {
                            attendanceRecord.StatusEO = null;
                            attendanceRecord.EarlyOut = null;
                            attendanceRecord.Remarks.Replace("[EO]", "");
                        }
                    }
                    else
                    {
                        attendanceRecord.StatusEO = null;
                        attendanceRecord.EarlyOut = null;
                        attendanceRecord.Remarks.Replace("[EO]", "");
                    }
                    //Calculate Late Out, Compare margin with Shift Late Out
                    if (attendanceRecord.TimeOut.Value.TimeOfDay > shiftEnd)
                    {
                        TimeSpan LateOutMinsSpan = (TimeSpan)(attendanceRecord.TimeOut.Value.TimeOfDay - shiftEnd);
                        if (LateOutMinsSpan.TotalMinutes > _shift.LateOut)
                        {
                            attendanceRecord.LateOut = (short)LateOutMinsSpan.TotalMinutes;
                            // Late Out cannot have an early out, In case of poll at multiple times before and after shiftend
                            attendanceRecord.EarlyOut = null;
                            attendanceRecord.StatusLO = true;
                            attendanceRecord.Remarks = attendanceRecord.Remarks + "[LO]";
                        }
                        else
                        {
                            attendanceRecord.LateOut = null;
                            attendanceRecord.LateOut = null;
                            attendanceRecord.Remarks.Replace("[LO]", "");
                        }
                    }
                    else
                    {
                        attendanceRecord.LateOut = null;
                        attendanceRecord.LateOut = null;
                        attendanceRecord.Remarks.Replace("[LO]", "");
                    }
                    attendanceRecord.WorkMin = (short)(mins.TotalMinutes);
                    if (attendanceRecord.EarlyIn != null && attendanceRecord.EarlyIn > _shift.EarlyIn)
                    {
                        attendanceRecord.WorkMin = (short)(attendanceRecord.WorkMin - attendanceRecord.EarlyIn);
                    }
                    if (attendanceRecord.LateOut != null && attendanceRecord.LateOut > _shift.LateOut)
                    {
                        attendanceRecord.WorkMin = (short)(attendanceRecord.WorkMin - attendanceRecord.LateOut);
                    }
                    if (attendanceRecord.EarlyIn == null && attendanceRecord.LateOut == null)
                    {

                    }
                    //round off work minutes
                    if (attendanceRecord.WorkMin > CalculateShiftMinutes(_shift, attendanceRecord.AttDate.Value.DayOfWeek) && (attendanceRecord.WorkMin <= (CalculateShiftMinutes(_shift, attendanceRecord.AttDate.Value.DayOfWeek) + _shift.OverTimeMin)))
                    {
                        attendanceRecord.WorkMin = CalculateShiftMinutes(_shift, attendanceRecord.AttDate.Value.DayOfWeek);
                    }
                    // RoundOff Overtime
                    if ((employee.EmpType.CatID == 2 || employee.EmpType.CatID == 4) && employee.CompanyID == 1)
                    {
                        if (attendanceRecord.OTMin > 0)
                        {
                            float OTmins = (float)attendanceRecord.OTMin;
                            float remainder = OTmins / 60;
                            int intpart = (int)remainder;
                            double fracpart = remainder - intpart;
                            if (fracpart < 0.5)
                            {
                                attendanceRecord.OTMin = (short)(intpart * 60);
                            }
                        }
                    }
                    ////Calculate OverTime, Compare margin with Shift OverTime
                    //if (attendanceRecord.EarlyIn > _shift.EarlyIn || attendanceRecord.LateOut > _shift.LateOut)
                    //{
                    //    if (attendanceRecord.StatusGZ != true || attendanceRecord.StatusDO != true)
                    //    {
                    //        short _EarlyIn;
                    //        short _LateOut;
                    //        if (attendanceRecord.EarlyIn == null)
                    //            _EarlyIn = 0;
                    //        else
                    //            _EarlyIn = 0;

                    //        if (attendanceRecord.LateOut == null)
                    //            _LateOut = 0;
                    //        else
                    //            _LateOut = (short)attendanceRecord.LateOut;

                    //        attendanceRecord.OTMin = (short)(_EarlyIn + _LateOut);
                    //        attendanceRecord.StatusOT = true;
                    //    }
                    //}
                    if ((attendanceRecord.StatusGZ != true || attendanceRecord.StatusDO != true) && employee.HasOT == true)
                    {
                        if (attendanceRecord.LateOut != null)
                        {
                            attendanceRecord.OTMin = attendanceRecord.LateOut;
                            attendanceRecord.StatusOT = true;
                            attendanceRecord.Remarks = attendanceRecord.Remarks + "[N-OT]";
                        }
                    }
                    //Mark Absent if less than 4 hours
                    if (attendanceRecord.AttDate.Value.DayOfWeek != DayOfWeek.Friday && attendanceRecord.StatusDO != true && attendanceRecord.StatusGZ != true)
                    {
                        short MinShiftMin = (short)_shift.MinHrs;
                        if (attendanceRecord.WorkMin < MinShiftMin)
                        {
                            attendanceRecord.StatusAB = true;
                            attendanceRecord.StatusP = false;
                            attendanceRecord.Remarks = attendanceRecord.Remarks + "[Absent]";
                        }
                        else
                        {
                            attendanceRecord.StatusAB = false;
                            attendanceRecord.StatusP = true;
                            attendanceRecord.Remarks.Replace("[Absent]", "");
                        }

                    }
                }


            }
            catch (Exception ex)
            {
                string _error = "";
                if (ex.InnerException.Message != null)
                    _error = ex.InnerException.Message;
                else
                    _error = ex.Message;
                _myHelperClass.WriteToLogFile("Attendance Processing Roster Times" + _error);
            }
        }
        #endregion

        #region -- Helper Function--
        private string ReturnDayOfWeek(DayOfWeek dayOfWeek)
        {
            string _DayName = "";
            switch (dayOfWeek)
            {
                case DayOfWeek.Monday:
                    _DayName = "Monday";
                    break;
                case DayOfWeek.Tuesday:
                    _DayName = "Tuesday";
                    break;
                case DayOfWeek.Wednesday:
                    _DayName = "Wednesday";
                    break;
                case DayOfWeek.Thursday:
                    _DayName = "Thursday";
                    break;
                case DayOfWeek.Friday:
                    _DayName = "Friday";
                    break;
                case DayOfWeek.Saturday:
                    _DayName = "Saturday";
                    break;
                case DayOfWeek.Sunday:
                    _DayName = "Sunday";
                    break;
            }
            return _DayName;
        }

        private TimeSpan CalculateShiftEndTime(Shift shift, DayOfWeek dayOfWeek)
        {
            Int16 workMins = 0;
            try
            {
                switch (dayOfWeek)
                {
                    case DayOfWeek.Monday:
                        workMins = shift.MonMin;
                        break;
                    case DayOfWeek.Tuesday:
                        workMins = shift.TueMin;
                        break;
                    case DayOfWeek.Wednesday:
                        workMins = shift.WedMin;
                        break;
                    case DayOfWeek.Thursday:
                        workMins = shift.ThuMin;
                        break;
                    case DayOfWeek.Friday:
                        workMins = shift.FriMin;
                        break;
                    case DayOfWeek.Saturday:
                        workMins = shift.SatMin;
                        break;
                    case DayOfWeek.Sunday:
                        workMins = shift.SunMin;
                        break;
                }
            }
            catch (Exception ex)
            {

            }
            return shift.StartTime + (new TimeSpan(0, workMins, 0));
        }

        private Int16 CalculateShiftMinutes(Shift shift, DayOfWeek dayOfWeek)
        {
            Int16 workMins = 0;
            try
            {
                switch (dayOfWeek)
                {
                    case DayOfWeek.Monday:
                        workMins = shift.MonMin;
                        break;
                    case DayOfWeek.Tuesday:
                        workMins = shift.TueMin;
                        break;
                    case DayOfWeek.Wednesday:
                        workMins = shift.WedMin;
                        break;
                    case DayOfWeek.Thursday:
                        workMins = shift.ThuMin;
                        break;
                    case DayOfWeek.Friday:
                        workMins = shift.FriMin;
                        break;
                    case DayOfWeek.Saturday:
                        workMins = shift.SatMin;
                        break;
                    case DayOfWeek.Sunday:
                        workMins = shift.SunMin;
                        break;
                }
            }
            catch (Exception ex)
            {

            }
            return workMins;
        }

        private DateTime CalculateShiftEndTime(Shift shift, DateTime _AttDate, TimeSpan _DutyTime)
        {
            Int16 workMins = 0;
            try
            {
                switch (_AttDate.Date.DayOfWeek)
                {
                    case DayOfWeek.Monday:
                        workMins = shift.MonMin;
                        break;
                    case DayOfWeek.Tuesday:
                        workMins = shift.TueMin;
                        break;
                    case DayOfWeek.Wednesday:
                        workMins = shift.WedMin;
                        break;
                    case DayOfWeek.Thursday:
                        workMins = shift.ThuMin;
                        break;
                    case DayOfWeek.Friday:
                        workMins = shift.FriMin;
                        break;
                    case DayOfWeek.Saturday:
                        workMins = shift.SatMin;
                        break;
                    case DayOfWeek.Sunday:
                        workMins = shift.SunMin;
                        break;
                }
            }
            catch (Exception ex)
            {

            }
            DateTime _datetime = new DateTime();
            TimeSpan _Time = new TimeSpan(0, workMins, 0);
            _datetime = _AttDate.Date.Add(_DutyTime);
            _datetime = _datetime.Add(_Time);
            return _datetime;
        }

        #endregion

        public void BootstrapAttendance(DateTime dateTime, List<Emp> emps, List<AttData> _AttData)
        {
            using (var ctx = new TAS2013Entities())
            {
                List<Roster> _Roster = new List<Roster>();
                _Roster = context.Rosters.Where(aa => aa.RosterDate == dateTime).ToList();
                List<RosterDetail> _NewRoster = new List<RosterDetail>();
                _NewRoster = context.RosterDetails.Where(aa => aa.RosterDate == dateTime).ToList();
                List<LvData> _LvData = new List<LvData>();
                _LvData = context.LvDatas.Where(aa => aa.AttDate == dateTime).ToList();
                List<LvShort> _lvShort = new List<LvShort>();
                _lvShort = context.LvShorts.Where(aa => aa.DutyDate == dateTime).ToList();
                //List<AttData> _AttData = context.AttDatas.Where(aa => aa.AttDate == dateTime).ToList();
                foreach (var emp in emps)
                {
                    string empDate = emp.EmpID + dateTime.ToString("yyMMdd");
                    if (_AttData.Where(aa => aa.EmpDate == empDate).Count() >0)
                    {
                        try
                        {
                            
                            /////////////////////////////////////////////////////
                            //  Mark Everyone Absent while creating Attendance //
                            /////////////////////////////////////////////////////
                            //Set DUTYCODE = D, StatusAB = true, and Remarks = [Absent]
                            AttData att = _AttData.First(aa => aa.EmpDate == empDate);
                            //Reset Flags
                            att.TimeIn = null;
                            att.TimeOut = null;
                            att.Tin0 = null;
                            att.Tout0 = null;
                            att.Tin1 = null;
                            att.Tout1 = null;
                            att.Tin2 = null;
                            att.Tout2 = null;
                            att.Tin3 = null;
                            att.Tout3 = null;
                            att.Tin4 = null;
                            att.Tout4 = null;
                            att.Tin5 = null;
                            att.Tout5 = null;
                            att.Tin6 = null;
                            att.Tout6 = null;
                            att.Tin7 = null;
                            att.Tout7 = null; 
                            att.Tin8 = null;
                            att.Tout8 = null;
                            att.StatusP = null;
                            att.StatusDO = null;
                            att.StatusEO = null;
                            att.StatusGZ = null;
                            att.StatusGZOT = null;
                            att.StatusHD = null;
                            att.StatusHL = null;
                            att.StatusIN = null;
                            att.StatusLeave = null;
                            att.StatusLI = null;
                            att.StatusLO = null;
                            att.StatusMN = null;
                            att.StatusOD = null;
                            att.StatusOT = null;
                            att.StatusSL = null;
                            att.WorkMin = null;
                            att.LateIn = null;
                            att.LateOut = null;
                            att.OTMin = null;
                            att.EarlyIn = null;
                            att.EarlyOut = null;
                            att.Remarks = null;
                            att.ShifMin = null;
                            att.SLMin = null;

                            att.AttDate = dateTime.Date;
                            att.DutyCode = "D";
                            att.StatusAB = true;
                            att.Remarks = "[Absent]";
                            if (emp.Shift != null)
                                att.DutyTime = emp.Shift.StartTime;
                            else
                                att.DutyTime = new TimeSpan(07, 45, 00);
                            att.EmpID = emp.EmpID;
                            att.EmpNo = emp.EmpNo;
                            att.EmpDate = emp.EmpID + dateTime.ToString("yyMMdd");
                            att.ShifMin = CalculateShiftMinutes(emp.Shift, dateTime.DayOfWeek);
                            //////////////////////////
                            //  Check for Rest Day //
                            ////////////////////////
                            //Set DutyCode = R, StatusAB=false, StatusDO = true, and Remarks=[DO]
                            //Check for 1st Day Off of Shift
                            if (emp.Shift.DaysName.Name == ReturnDayOfWeek(dateTime.DayOfWeek))
                            {
                                att.DutyCode = "R";
                                att.StatusAB = false;
                                att.StatusDO = true;
                                att.Remarks = "[DO]";
                            }
                            //Check for 2nd Day Off of shift
                            if (emp.Shift.DaysName1.Name == ReturnDayOfWeek(dateTime.DayOfWeek))
                            {
                                att.DutyCode = "R";
                                att.StatusAB = false;
                                att.StatusDO = true;
                                att.Remarks = "[DO]";
                            }
                            //////////////////////////
                            //  Check for Roster   //
                            ////////////////////////
                            //If Roster DutyCode is Rest then change the StatusAB and StatusDO
                            foreach (var roster in _Roster.Where(aa => aa.EmpDate == att.EmpDate))
                            {
                                att.DutyCode = roster.DutyCode.Trim();
                                if (att.DutyCode == "R")
                                {
                                    att.StatusAB = false;
                                    att.StatusDO = true;
                                    att.DutyCode = "R";
                                    att.Remarks = "[DO]";
                                }
                                att.ShifMin = roster.WorkMin;
                                att.DutyTime = roster.DutyTime;
                            }

                            ////New Roster
                            string empCdate = "Emp" + emp.EmpID.ToString() + dateTime.ToString("yyMMdd");
                            string sectionCdate = "Section" + emp.SecID.ToString() + dateTime.ToString("yyMMdd");
                            string crewCdate = "Crew" + emp.CrewID.ToString() + dateTime.ToString("yyMMdd");
                            string shiftCdate = "Shift" + emp.ShiftID.ToString() + dateTime.ToString("yyMMdd");
                            if (_NewRoster.Where(aa => aa.CriteriaValueDate == empCdate).Count() > 0)
                            {
                                var roster = _NewRoster.FirstOrDefault(aa => aa.CriteriaValueDate == empCdate);
                                if (roster.WorkMin == 0)
                                {
                                    att.StatusAB = false;
                                    att.StatusDO = true;
                                    att.Remarks = "[DO]";
                                    att.DutyCode = "R";
                                    att.ShifMin = 0;
                                }
                                else
                                {
                                    att.ShifMin = roster.WorkMin;
                                    att.DutyCode = "D";
                                    att.DutyTime = roster.DutyTime;
                                }
                            }
                            else if (_NewRoster.Where(aa => aa.CriteriaValueDate == sectionCdate).Count() > 0)
                            {
                                var roster = _NewRoster.FirstOrDefault(aa => aa.CriteriaValueDate == sectionCdate);
                                if (roster.WorkMin == 0)
                                {
                                    att.StatusAB = false;
                                    att.StatusDO = true;
                                    att.Remarks = "[DO]";
                                    att.DutyCode = "R";
                                    att.ShifMin = 0;
                                }
                                else
                                {
                                    att.ShifMin = roster.WorkMin;
                                    att.DutyCode = "D";
                                    att.DutyTime = roster.DutyTime;
                                }
                            }
                            else if (_NewRoster.Where(aa => aa.CriteriaValueDate == crewCdate).Count() > 0)
                            {
                                var roster = _NewRoster.FirstOrDefault(aa => aa.CriteriaValueDate == crewCdate);
                                if (roster.WorkMin == 0)
                                {
                                    att.StatusAB = false;
                                    att.StatusDO = true;
                                    att.Remarks = "[DO]";
                                    att.DutyCode = "R";
                                    att.ShifMin = 0;
                                }
                                else
                                {
                                    att.ShifMin = roster.WorkMin;
                                    att.DutyCode = "D";
                                    att.DutyTime = roster.DutyTime;
                                }
                            }
                            else if (_NewRoster.Where(aa => aa.CriteriaValueDate == shiftCdate).Count() > 0)
                            {
                                var roster = _NewRoster.FirstOrDefault(aa => aa.CriteriaValueDate == shiftCdate);
                                if (roster.WorkMin == 0)
                                {
                                    att.StatusAB = false;
                                    att.StatusDO = true;
                                    att.Remarks = "[DO]";
                                    att.DutyCode = "R";
                                    att.ShifMin = 0;
                                }
                                else
                                {
                                    att.ShifMin = roster.WorkMin;
                                    att.DutyCode = "D";
                                    att.DutyTime = roster.DutyTime;
                                }
                            }

                            //////////////////////////
                            //  Check for GZ Day //
                            ////////////////////////
                            //Set DutyCode = R, StatusAB=false, StatusGZ = true, and Remarks=[GZ]
                            if (emp.Shift.GZDays == true)
                            {
                                foreach (var holiday in context.Holidays)
                                {
                                    if (context.Holidays.Where(hol => hol.HolDate.Month == att.AttDate.Value.Month && hol.HolDate.Day == att.AttDate.Value.Day).Count() > 0)
                                    {
                                        att.DutyCode = "G";
                                        att.StatusAB = false;
                                        att.StatusGZ = true;
                                        att.Remarks = "[GZ]";
                                        att.ShifMin = 0;
                                    }
                                }
                            }
                            ////////////////////////////
                            //TODO Check for Job Card//
                            //////////////////////////



                            ////////////////////////////
                            //  Check for Short Leave//
                            //////////////////////////
                            foreach (var sLeave in _lvShort.Where(aa => aa.EmpDate == att.EmpDate))
                            {
                                if (_lvShort.Where(lv => lv.EmpDate == att.EmpDate).Count() > 0)
                                {
                                    att.StatusSL = true;
                                    att.StatusAB = null;
                                    att.DutyCode = "L";
                                    att.Remarks = "[Short Leave]";
                                }
                            }

                            //////////////////////////
                            //   Check for Leave   //
                            ////////////////////////
                            //Set DutyCode = R, StatusAB=false, StatusGZ = true, and Remarks=[GZ]
                            foreach (var Leave in _LvData)
                            {
                                var _Leave = _LvData.Where(lv => lv.EmpDate == att.EmpDate && lv.HalfLeave != true);
                                if (_Leave.Count() > 0)
                                {
                                    att.StatusLeave = true;
                                    att.StatusAB = false;
                                    att.DutyCode = "L";
                                    att.StatusDO = false;
                                    if (Leave.LvCode == "A")
                                        att.Remarks = "[CL]";
                                    else if (Leave.LvCode == "B")
                                        att.Remarks = "[AL]";
                                    else if (Leave.LvCode == "C")
                                        att.Remarks = "[SL]";
                                    else
                                        att.Remarks = "[" + _Leave.FirstOrDefault().LvType.LvDesc + "]";
                                }
                                else
                                {
                                    att.StatusLeave = false;
                                }
                            }

                             //////////////////////////
                            //Check for Half Leave///
                           /////////////////////////
                            var _HalfLeave = _LvData.Where(lv => lv.EmpDate == att.EmpDate && lv.HalfLeave == true);
                            if (_HalfLeave.Count() > 0)
                            {
                                att.StatusLeave = true;
                                att.StatusAB = false;
                                att.DutyCode = "L";
                                att.StatusHL = true;
                                att.StatusDO = false;
                                if (_HalfLeave.FirstOrDefault().LvCode == "A")
                                    att.Remarks = "[H-CL]";
                                else if (_HalfLeave.FirstOrDefault().LvCode == "B")
                                    att.Remarks = "[S-AL]";
                                else if (_HalfLeave.FirstOrDefault().LvCode == "C")
                                    att.Remarks = "[H-SL]";
                                else
                                    att.Remarks = "[Half Leave]";
                            }
                            else
                            {
                                att.StatusLeave = false;
                            }
                            ctx.SaveChanges();
                        }
                        catch (Exception ex)
                        {
                            _myHelperClass.WriteToLogFile("-------Error In Creating Attendance of Employee: " + emp.EmpNo + " ------" + ex.InnerException.Message);
                        }
                    }
                }
                ctx.Dispose();
            }
        }

        #endregion
    }
}
