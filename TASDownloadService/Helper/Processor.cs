using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using TASDownloadService.Model;

namespace TASDownloadService.Helper
{
    public class Processor
    {
        //TAS2013Entities context = new TAS2013Entities();
        //public void ProcessAttendance()
        //{
        //    //create attendance first, compare with attprocess table
        //    if (context.AttProcesses.Where(proc => proc.ProcessDate == DateTime.Today).Count() == 0)
        //    {
        //        CreateAttendance(DateTime.Today.Date);
        //    }
        //    DateTime dt = new DateTime();
        //    dt = DateTime.Today.Date.AddDays(-4);
        //    List<PollData> unprocessedPolls = context.PollDatas.Where(p => p.Process == false).ToList();
        //    foreach (PollData up in unprocessedPolls)
        //    {
        //        try
        //        {
        //            //Create Attendance if any poll date is not processed already
        //            if (context.AttProcesses.Where(ap => ap.ProcessDate == up.EntDate).Count() == 0)
        //            {
        //                CreateAttendance(up.EntDate.Date);
        //            }
        //            //Check AttData with EmpDate
        //            if (context.AttDatas.Where(attd => attd.EmpDate == up.EmpDate).Count() > 0)
        //            {
        //                AttData attendanceRecord = context.AttDatas.First(attd => attd.EmpDate == up.EmpDate);
        //                //Set Time In and Time Out
        //                CalculateTimeINOUT(attendanceRecord, up);
        //                Emp employee = attendanceRecord.Emp;
        //                Shift shift = employee.Shift;
        //                //If TimeIn and TimeOut are not null, then calculate other Atributes
        //                if (context.AttDatas.First(attd => attd.EmpDate == up.EmpDate).TimeIn != null && context.AttDatas.First(attd => attd.EmpDate == up.EmpDate).TimeOut != null)
        //                {
        //                    if (context.Rosters.Where(r => r.EmpDate == up.EmpDate).Count() > 0)
        //                    {
        //                        CalculateRosterTimes(attendanceRecord, context.Rosters.FirstOrDefault(r => r.EmpDate == up.EmpDate));
        //                    }
        //                    else
        //                    {
        //                        CalculateShiftTimes(attendanceRecord, shift);
        //                    }
        //                }

        //            }
        //            up.Process = true;
        //        }
        //        catch (Exception ex)
        //        {

        //        }
        //        context.SaveChanges();
        //    }

        //}

        //private void CalculateTimeINOUT(AttData attendanceRecord, PollData up)
        //{
        //    try
        //    {
        //        switch (up.RdrDuty)
        //        {
        //            case 1: //IN
        //                if (attendanceRecord.Tin0 == null)
        //                {
        //                    attendanceRecord.Tin0 = up.EntTime.TimeOfDay;
        //                    attendanceRecord.TimeIn = up.EntTime.TimeOfDay;
        //                    attendanceRecord.StatusAB = false;
        //                    attendanceRecord.StatusP = true;
        //                    attendanceRecord.Remarks = null;
        //                }
        //                else if (attendanceRecord.Tin1 == null)
        //                {
        //                    attendanceRecord.Tin1 = up.EntTime.TimeOfDay;
        //                    SortingInOutTime(attendanceRecord);
        //                }
        //                else if (attendanceRecord.Tin2 == null)
        //                {
        //                    attendanceRecord.Tin2 = up.EntTime.TimeOfDay;
        //                    SortingInOutTime(attendanceRecord);
        //                }
        //                else if (attendanceRecord.Tin3 == null)
        //                {
        //                    attendanceRecord.Tin3 = up.EntTime.TimeOfDay;
        //                    SortingInOutTime(attendanceRecord);
        //                }
        //                else if (attendanceRecord.Tin4 == null)
        //                {
        //                    attendanceRecord.Tin4 = up.EntTime.TimeOfDay;
        //                    SortingInOutTime(attendanceRecord);
        //                }
        //                else if (attendanceRecord.Tin5 == null)
        //                {
        //                    attendanceRecord.Tin5 = up.EntTime.TimeOfDay;
        //                    SortingInOutTime(attendanceRecord);
        //                }
        //                else if (attendanceRecord.Tin6 == null)
        //                {
        //                    attendanceRecord.Tin6 = up.EntTime.TimeOfDay;
        //                    SortingInOutTime(attendanceRecord);
        //                }
        //                else if (attendanceRecord.Tin7 == null)
        //                {
        //                    attendanceRecord.Tin7 = up.EntTime.TimeOfDay;
        //                    SortingInOutTime(attendanceRecord);
        //                }
        //                else if (attendanceRecord.Tin8 == null)
        //                {
        //                    attendanceRecord.Tin8 = up.EntTime.TimeOfDay;
        //                    SortingInOutTime(attendanceRecord);
        //                }
        //                else if (attendanceRecord.Tin9 == null)
        //                {
        //                    attendanceRecord.Tin9 = up.EntTime.TimeOfDay;
        //                    SortingInOutTime(attendanceRecord);
        //                }
        //                else if (attendanceRecord.Tin10 == null)
        //                {
        //                    attendanceRecord.Tin10 = up.EntTime.TimeOfDay;
        //                    SortingInOutTime(attendanceRecord);
        //                }
        //                else
        //                {
        //                    attendanceRecord.Tin11 = up.EntTime.TimeOfDay;
        //                    SortingInOutTime(attendanceRecord);
        //                }
        //                break;
        //            case 5: //OUT
        //                if (attendanceRecord.Tout0 == null)
        //                {
        //                    attendanceRecord.Tout0 = up.EntTime.TimeOfDay;
        //                    attendanceRecord.TimeOut = up.EntTime.TimeOfDay;
        //                    SortingOutTime(attendanceRecord);
        //                }
        //                else if (attendanceRecord.Tout1 == null)
        //                {
        //                    attendanceRecord.Tout1 = up.EntTime.TimeOfDay;
        //                    attendanceRecord.TimeOut = up.EntTime.TimeOfDay;
        //                    SortingOutTime(attendanceRecord);
        //                }
        //                else if (attendanceRecord.Tout2 == null)
        //                {
        //                    attendanceRecord.Tout2 = up.EntTime.TimeOfDay;
        //                    attendanceRecord.TimeOut = up.EntTime.TimeOfDay;
        //                    SortingOutTime(attendanceRecord);
        //                }
        //                else if (attendanceRecord.Tout3 == null)
        //                {
        //                    attendanceRecord.Tout3 = up.EntTime.TimeOfDay;
        //                    attendanceRecord.TimeOut = up.EntTime.TimeOfDay;
        //                    SortingOutTime(attendanceRecord);
        //                }
        //                else if (attendanceRecord.Tout4 == null)
        //                {
        //                    attendanceRecord.Tout4 = up.EntTime.TimeOfDay;
        //                    attendanceRecord.TimeOut = up.EntTime.TimeOfDay;
        //                    SortingOutTime(attendanceRecord);
        //                }
        //                else if (attendanceRecord.Tout5 == null)
        //                {
        //                    attendanceRecord.Tout5 = up.EntTime.TimeOfDay;
        //                    attendanceRecord.TimeOut = up.EntTime.TimeOfDay;
        //                    SortingOutTime(attendanceRecord);
        //                }
        //                else if (attendanceRecord.Tout6 == null)
        //                {
        //                    attendanceRecord.Tout6 = up.EntTime.TimeOfDay;
        //                    attendanceRecord.TimeOut = up.EntTime.TimeOfDay;
        //                    SortingOutTime(attendanceRecord);
        //                }
        //                else if (attendanceRecord.Tout7 == null)
        //                {
        //                    attendanceRecord.Tout7 = up.EntTime.TimeOfDay;
        //                    attendanceRecord.TimeOut = up.EntTime.TimeOfDay;
        //                    SortingOutTime(attendanceRecord);
        //                }
        //                else if (attendanceRecord.Tout8 == null)
        //                {
        //                    attendanceRecord.Tout8 = up.EntTime.TimeOfDay;
        //                    attendanceRecord.TimeOut = up.EntTime.TimeOfDay;
        //                    SortingOutTime(attendanceRecord);
        //                }
        //                else if (attendanceRecord.Tout9 == null)
        //                {
        //                    attendanceRecord.Tout9 = up.EntTime.TimeOfDay;
        //                    attendanceRecord.TimeOut = up.EntTime.TimeOfDay;
        //                    SortingOutTime(attendanceRecord);
        //                }
        //                else if (attendanceRecord.Tout10 == null)
        //                {
        //                    attendanceRecord.Tout10 = up.EntTime.TimeOfDay;
        //                    attendanceRecord.TimeOut = up.EntTime.TimeOfDay;
        //                    SortingOutTime(attendanceRecord);
        //                }
        //                else
        //                {
        //                    attendanceRecord.Tout11 = up.EntTime.TimeOfDay;
        //                    attendanceRecord.TimeOut = up.EntTime.TimeOfDay;
        //                    SortingOutTime(attendanceRecord);
        //                }
        //                break;
        //            case 8: //DUTY
        //                if (attendanceRecord.Tin0 != null)
        //                {
        //                    if (attendanceRecord.Tout0 == null)
        //                    {
        //                        attendanceRecord.Tout0 = up.EntTime.TimeOfDay;
        //                        attendanceRecord.TimeOut = up.EntTime.TimeOfDay;
        //                    }
        //                    else if (attendanceRecord.Tin1 == null)
        //                    {
        //                        attendanceRecord.Tin1 = up.EntTime.TimeOfDay;
        //                        attendanceRecord.TimeOut = up.EntTime.TimeOfDay;
        //                    }
        //                    else if (attendanceRecord.Tout1 == null)
        //                    {
        //                        attendanceRecord.Tout1 = up.EntTime.TimeOfDay;
        //                        attendanceRecord.TimeOut = up.EntTime.TimeOfDay;
        //                    }
        //                    else if (attendanceRecord.Tin2 == null)
        //                    {
        //                        attendanceRecord.Tin2 = up.EntTime.TimeOfDay;
        //                        attendanceRecord.TimeOut = up.EntTime.TimeOfDay;
        //                    }
        //                    else if (attendanceRecord.Tout2 == null)
        //                    {
        //                        attendanceRecord.Tout2 = up.EntTime.TimeOfDay;
        //                        attendanceRecord.TimeOut = up.EntTime.TimeOfDay;
        //                    }
        //                    else if (attendanceRecord.Tin3 == null)
        //                    {
        //                        attendanceRecord.Tin3 = up.EntTime.TimeOfDay;
        //                        attendanceRecord.TimeOut = up.EntTime.TimeOfDay;
        //                    }
        //                    else if (attendanceRecord.Tout3 == null)
        //                    {
        //                        attendanceRecord.Tout3 = up.EntTime.TimeOfDay;
        //                        attendanceRecord.TimeOut = up.EntTime.TimeOfDay;
        //                    }
        //                    else if (attendanceRecord.Tin4 == null)
        //                    {
        //                        attendanceRecord.Tin4 = up.EntTime.TimeOfDay;
        //                        attendanceRecord.TimeOut = up.EntTime.TimeOfDay;
        //                    }
        //                    else if (attendanceRecord.Tout4 == null)
        //                    {
        //                        attendanceRecord.Tout4 = up.EntTime.TimeOfDay;
        //                        attendanceRecord.TimeOut = up.EntTime.TimeOfDay;
        //                    }
        //                    else if (attendanceRecord.Tin5 == null)
        //                    {
        //                        attendanceRecord.Tin5 = up.EntTime.TimeOfDay;
        //                        attendanceRecord.TimeOut = up.EntTime.TimeOfDay;
        //                    }
        //                    else if (attendanceRecord.Tout5 == null)
        //                    {
        //                        attendanceRecord.Tout5 = up.EntTime.TimeOfDay;
        //                        attendanceRecord.TimeOut = up.EntTime.TimeOfDay;
        //                    }
        //                    else if (attendanceRecord.Tin6 == null)
        //                    {
        //                        attendanceRecord.Tin6 = up.EntTime.TimeOfDay;
        //                        attendanceRecord.TimeOut = up.EntTime.TimeOfDay;
        //                    }
        //                    else if (attendanceRecord.Tout6 == null)
        //                    {
        //                        attendanceRecord.Tout6 = up.EntTime.TimeOfDay;
        //                        attendanceRecord.TimeOut = up.EntTime.TimeOfDay;
        //                    }
        //                    //
        //                    else if (attendanceRecord.Tin7 == null)
        //                    {
        //                        attendanceRecord.Tin7 = up.EntTime.TimeOfDay;
        //                        attendanceRecord.TimeOut = up.EntTime.TimeOfDay;
        //                    }
        //                    else if (attendanceRecord.Tout7 == null)
        //                    {
        //                        attendanceRecord.Tout7 = up.EntTime.TimeOfDay;
        //                        attendanceRecord.TimeOut = up.EntTime.TimeOfDay;
        //                    }
        //                    else if (attendanceRecord.Tin8 == null)
        //                    {
        //                        attendanceRecord.Tin8 = up.EntTime.TimeOfDay;
        //                        attendanceRecord.TimeOut = up.EntTime.TimeOfDay;
        //                    }
        //                    else if (attendanceRecord.Tout8 == null)
        //                    {
        //                        attendanceRecord.Tout8 = up.EntTime.TimeOfDay;
        //                        attendanceRecord.TimeOut = up.EntTime.TimeOfDay;
        //                    }
        //                    else if (attendanceRecord.Tin9 == null)
        //                    {
        //                        attendanceRecord.Tin9 = up.EntTime.TimeOfDay;
        //                        attendanceRecord.TimeOut = up.EntTime.TimeOfDay;
        //                    }
        //                    else if (attendanceRecord.Tout9 == null)
        //                    {
        //                        attendanceRecord.Tout9 = up.EntTime.TimeOfDay;
        //                        attendanceRecord.TimeOut = up.EntTime.TimeOfDay;
        //                    }
        //                    else if (attendanceRecord.Tin10 == null)
        //                    {
        //                        attendanceRecord.Tin10 = up.EntTime.TimeOfDay;
        //                        attendanceRecord.TimeOut = up.EntTime.TimeOfDay;
        //                    }
        //                    else if (attendanceRecord.Tout10 == null)
        //                    {
        //                        attendanceRecord.Tout10 = up.EntTime.TimeOfDay;
        //                        attendanceRecord.TimeOut = up.EntTime.TimeOfDay;
        //                    }
        //                    else if (attendanceRecord.Tin11 == null)
        //                    {
        //                        attendanceRecord.Tin11 = up.EntTime.TimeOfDay;
        //                        attendanceRecord.TimeOut = up.EntTime.TimeOfDay;
        //                    }
        //                    else if (attendanceRecord.Tout11 == null)
        //                    {
        //                        attendanceRecord.Tout11 = up.EntTime.TimeOfDay;
        //                        attendanceRecord.TimeOut = up.EntTime.TimeOfDay;
        //                    }
        //                    else if (attendanceRecord.Tin12 == null)
        //                    {
        //                        attendanceRecord.Tin12 = up.EntTime.TimeOfDay;
        //                        attendanceRecord.TimeOut = up.EntTime.TimeOfDay;
        //                    }
        //                    else if (attendanceRecord.Tout12 == null)
        //                    {
        //                        attendanceRecord.Tout12 = up.EntTime.TimeOfDay;
        //                        attendanceRecord.TimeOut = up.EntTime.TimeOfDay;
        //                    }
        //                    else if (attendanceRecord.Tin13 == null)
        //                    {
        //                        attendanceRecord.Tin13 = up.EntTime.TimeOfDay;
        //                        attendanceRecord.TimeOut = up.EntTime.TimeOfDay;
        //                    }
        //                    else
        //                    {
        //                        attendanceRecord.Tout13 = up.EntTime.TimeOfDay;
        //                        attendanceRecord.TimeOut = up.EntTime.TimeOfDay;
        //                    }
        //                }
        //                else
        //                {
        //                    attendanceRecord.Tin0 = up.EntTime.TimeOfDay;
        //                    attendanceRecord.TimeIn = up.EntTime.TimeOfDay;
        //                    attendanceRecord.StatusAB = false;
        //                    attendanceRecord.StatusP = true;
        //                    attendanceRecord.Remarks = null;
        //                }
        //                break;
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        //Error in TimeIN/OUT
        //    }
        //}

        //// Sorting Time In
        //private void SortingInOutTime(AttData attendanceRecord)
        //{
        //    List<TimeSpan> _InTimes = new List<TimeSpan>();

        //    if (attendanceRecord.Tin0 != null)
        //        _InTimes.Add((TimeSpan)attendanceRecord.Tin0);
        //    if (attendanceRecord.Tin1 != null)
        //        _InTimes.Add((TimeSpan)attendanceRecord.Tin1);
        //    if (attendanceRecord.Tin2 != null)
        //        _InTimes.Add((TimeSpan)attendanceRecord.Tin2);
        //    if (attendanceRecord.Tin3 != null)
        //        _InTimes.Add((TimeSpan)attendanceRecord.Tin3);
        //    if (attendanceRecord.Tin4 != null)
        //        _InTimes.Add((TimeSpan)attendanceRecord.Tin4);
        //    if (attendanceRecord.Tin5 != null)
        //        _InTimes.Add((TimeSpan)attendanceRecord.Tin5);
        //    if (attendanceRecord.Tin6 != null)
        //        _InTimes.Add((TimeSpan)attendanceRecord.Tin6);
        //    if (attendanceRecord.Tin7 != null)
        //        _InTimes.Add((TimeSpan)attendanceRecord.Tin7);
        //    if (attendanceRecord.Tin8 != null)
        //        _InTimes.Add((TimeSpan)attendanceRecord.Tin8);
        //    if (attendanceRecord.Tin9 != null)
        //        _InTimes.Add((TimeSpan)attendanceRecord.Tin9);
        //    if (attendanceRecord.Tin10 != null)
        //        _InTimes.Add((TimeSpan)attendanceRecord.Tin10);

        //    var list = _InTimes.OrderBy(x => x.Hours).ToList();
        //    PlacedSortedInTime(attendanceRecord, list);

        //}
        //private void PlacedSortedInTime(AttData attendanceRecord, List<TimeSpan> _InTimes)
        //{
        //    for (int i = 0; i < _InTimes.Count; i++)
        //    {
        //        switch (i)
        //        {
        //            case 0:
        //                attendanceRecord.Tin0 = _InTimes[i];
        //                attendanceRecord.TimeIn = _InTimes[i];
        //                break;
        //            case 1:
        //                attendanceRecord.Tin1 = _InTimes[i];
        //                break;
        //            case 2:
        //                attendanceRecord.Tin2 = _InTimes[i];
        //                break;
        //            case 3:
        //                attendanceRecord.Tin3 = _InTimes[i];
        //                break;
        //            case 4:
        //                attendanceRecord.Tin4 = _InTimes[i];
        //                break;
        //            case 5:
        //                attendanceRecord.Tin5 = _InTimes[i];
        //                break;
        //            case 6:
        //                attendanceRecord.Tin6 = _InTimes[i];
        //                break;
        //            case 7:
        //                attendanceRecord.Tin7 = _InTimes[i];
        //                break;
        //            case 8:
        //                attendanceRecord.Tin8 = _InTimes[i];
        //                break;
        //            case 9:
        //                attendanceRecord.Tin9 = _InTimes[i];
        //                break;
        //        }
        //    }
        //}
        ////Sorting Time Out
        //private void SortingOutTime(AttData attendanceRecord)
        //{
        //    List<TimeSpan> _OutTimes = new List<TimeSpan>();

        //    if (attendanceRecord.Tout0 != null)
        //        _OutTimes.Add((TimeSpan)attendanceRecord.Tout0);
        //    if (attendanceRecord.Tout1 != null)
        //        _OutTimes.Add((TimeSpan)attendanceRecord.Tout1);
        //    if (attendanceRecord.Tout2 != null)
        //        _OutTimes.Add((TimeSpan)attendanceRecord.Tout2);
        //    if (attendanceRecord.Tout3 != null)
        //        _OutTimes.Add((TimeSpan)attendanceRecord.Tout3);
        //    if (attendanceRecord.Tout4 != null)
        //        _OutTimes.Add((TimeSpan)attendanceRecord.Tout4);
        //    if (attendanceRecord.Tout5 != null)
        //        _OutTimes.Add((TimeSpan)attendanceRecord.Tout5);
        //    if (attendanceRecord.Tout6 != null)
        //        _OutTimes.Add((TimeSpan)attendanceRecord.Tout6);
        //    if (attendanceRecord.Tout7 != null)
        //        _OutTimes.Add((TimeSpan)attendanceRecord.Tout7);
        //    if (attendanceRecord.Tout8 != null)
        //        _OutTimes.Add((TimeSpan)attendanceRecord.Tout8);
        //    if (attendanceRecord.Tout9 != null)
        //        _OutTimes.Add((TimeSpan)attendanceRecord.Tout9);
        //    if (attendanceRecord.Tout10 != null)
        //        _OutTimes.Add((TimeSpan)attendanceRecord.Tout10);

        //    var list = _OutTimes.OrderBy(x => x.Hours).ToList();
        //    PlacedSortedOutTime(attendanceRecord, list);


        //}

        //private void PlacedSortedOutTime(AttData attendanceRecord, List<TimeSpan> _OutTimes)
        //{
        //    for (int i = 0; i < _OutTimes.Count; i++)
        //    {
        //        switch (i)
        //        {
        //            case 0:
        //                attendanceRecord.Tout0 = _OutTimes[i];
        //                attendanceRecord.TimeOut = _OutTimes[i];
        //                break;
        //            case 1:
        //                attendanceRecord.Tout1 = _OutTimes[i];
        //                attendanceRecord.TimeOut = _OutTimes[i];
        //                break;
        //            case 2:
        //                attendanceRecord.Tout2 = _OutTimes[i];
        //                attendanceRecord.TimeOut = _OutTimes[i];
        //                break;
        //            case 3:
        //                attendanceRecord.Tout3 = _OutTimes[i];
        //                attendanceRecord.TimeOut = _OutTimes[i];
        //                break;
        //            case 4:
        //                attendanceRecord.Tout4 = _OutTimes[i];
        //                attendanceRecord.TimeOut = _OutTimes[i];
        //                break;
        //            case 5:
        //                attendanceRecord.Tout5 = _OutTimes[i];
        //                attendanceRecord.TimeOut = _OutTimes[i];
        //                break;
        //            case 6:
        //                attendanceRecord.Tout6 = _OutTimes[i];
        //                attendanceRecord.TimeOut = _OutTimes[i];
        //                break;
        //            case 7:
        //                attendanceRecord.Tout7 = _OutTimes[i];
        //                attendanceRecord.TimeOut = _OutTimes[i];
        //                break;
        //            case 8:
        //                attendanceRecord.Tout8 = _OutTimes[i];
        //                attendanceRecord.TimeOut = _OutTimes[i];
        //                break;
        //            case 9:
        //                attendanceRecord.Tout9 = _OutTimes[i];
        //                attendanceRecord.TimeOut = _OutTimes[i];
        //                break;
        //        }
        //    }
        //}

        //private void CalculateShiftTimes(AttData attendanceRecord, Shift shift)
        //{
        //    try
        //    {
        //        //Calculate WorkMin
        //        TimeSpan mins = (TimeSpan)(attendanceRecord.TimeOut - attendanceRecord.TimeIn);
        //        //Check if GZ holiday then place all WorkMin in GZOTMin
        //        if (attendanceRecord.StatusGZ == true)
        //        {
        //            attendanceRecord.GZOTMin = (short)mins.TotalMinutes;
        //            attendanceRecord.WorkMin = (short)mins.TotalMinutes;
        //            attendanceRecord.StatusGZOT = true;
        //        }
        //        else if (attendanceRecord.StatusDO == true)
        //        {
        //            attendanceRecord.OTMin = (short)mins.TotalMinutes;
        //            attendanceRecord.WorkMin = (short)mins.TotalMinutes;
        //            attendanceRecord.StatusOT = true;
        //        }
        //        else
        //        {
        //            if (shift.HasBreak == true)
        //            {
        //                attendanceRecord.WorkMin = (short)(mins.TotalMinutes - shift.BreakMin);
        //                attendanceRecord.ShifMin = (short)(CalculateShiftMinutes(shift, attendanceRecord.AttDate.Value.DayOfWeek) - (short)shift.BreakMin);
        //            }
        //            else
        //            {
        //                //Calculate Late IN, Compare margin with Shift Late In
        //                if (attendanceRecord.TimeIn.Value > shift.StartTime)
        //                {
        //                    TimeSpan lateMinsSpan = (TimeSpan)(attendanceRecord.TimeIn - shift.StartTime);
        //                    if (lateMinsSpan.TotalMinutes > shift.LateIn)
        //                    {
        //                        attendanceRecord.LateIn = (short)lateMinsSpan.TotalMinutes;
        //                        attendanceRecord.StatusLI = true;
        //                        attendanceRecord.EarlyIn = null;
        //                    }
        //                }

        //                //Calculate Early In, Compare margin with Shift Early In
        //                if (attendanceRecord.TimeIn.Value < shift.StartTime)
        //                {
        //                    TimeSpan EarlyInMinsSpan = (TimeSpan)(shift.StartTime - attendanceRecord.TimeIn);
        //                    if (EarlyInMinsSpan.TotalMinutes > shift.EarlyIn)
        //                    {
        //                        attendanceRecord.EarlyIn = (short)EarlyInMinsSpan.TotalMinutes;
        //                        attendanceRecord.StatusEI = true;
        //                        attendanceRecord.LateIn = null;
        //                    }
        //                }

        //                // CalculateShiftEndTime = ShiftStart + DutyHours
        //                TimeSpan shiftEnd = CalculateShiftEndTime(shift, attendanceRecord.AttDate.Value.DayOfWeek);

        //                //Calculate Early Out, Compare margin with Shift Early Out
        //                if (attendanceRecord.TimeOut.Value < shiftEnd)
        //                {
        //                    TimeSpan EarlyOutMinsSpan = (TimeSpan)(shiftEnd - attendanceRecord.TimeOut);
        //                    if (EarlyOutMinsSpan.TotalMinutes > shift.EarlyOut)
        //                    {
        //                        attendanceRecord.EarlyOut = (short)EarlyOutMinsSpan.TotalMinutes;
        //                        attendanceRecord.StatusEO = true;
        //                        attendanceRecord.LateOut = null;
        //                    }
        //                }

        //                //Calculate Late Out, Compare margin with Shift Late Out
        //                if (attendanceRecord.TimeOut.Value > shiftEnd)
        //                {
        //                    TimeSpan LateOutMinsSpan = (TimeSpan)(attendanceRecord.TimeOut - shiftEnd);
        //                    if (LateOutMinsSpan.TotalMinutes > shift.LateOut)
        //                    {
        //                        attendanceRecord.LateOut = (short)LateOutMinsSpan.TotalMinutes;
        //                        // Late Out cannot have an early out, In case of poll at multiple times before and after shiftend
        //                        attendanceRecord.EarlyOut = null;
        //                        attendanceRecord.StatusLO = true;
        //                    }
        //                }
        //                //Subtract EarlyIn and LateOut from Work Minutes
        //                if (attendanceRecord.EarlyIn != null && attendanceRecord.EarlyIn > shift.EarlyIn)
        //                {
        //                    attendanceRecord.WorkMin = (short)(mins.TotalMinutes - attendanceRecord.EarlyIn);
        //                }
        //                else if (attendanceRecord.LateOut != null && attendanceRecord.LateOut > shift.LateOut)
        //                {
        //                    attendanceRecord.WorkMin = (short)(mins.TotalMinutes - attendanceRecord.LateOut);
        //                }
        //                else
        //                    attendanceRecord.WorkMin = (short)(mins.TotalMinutes);
        //                //Calculate OverTime, Compare margin with Shift OverTime
        //                if (attendanceRecord.EarlyIn > shift.EarlyIn || attendanceRecord.LateOut > shift.LateOut)
        //                {
        //                    if (attendanceRecord.StatusGZ != true || attendanceRecord.StatusDO != true)
        //                    {
        //                        short _EarlyIn;
        //                        short _LateOut;
        //                        if (attendanceRecord.EarlyIn == null)
        //                            _EarlyIn = 0;
        //                        else
        //                            _EarlyIn = (short)attendanceRecord.EarlyIn;

        //                        if (attendanceRecord.LateOut == null)
        //                            _LateOut = 0;
        //                        else
        //                            _LateOut = (short)attendanceRecord.LateOut;

        //                        attendanceRecord.OTMin = (short)(_EarlyIn + _LateOut);
        //                        attendanceRecord.StatusOT = true;
        //                    }
        //                }


        //            }
        //        }




        //    }
        //    catch (Exception ex)
        //    {

        //    }
        //}

        //private void CalculateRosterTimes(AttData attendanceRecord, Roster roster)
        //{
        //    try
        //    {
        //        TimeSpan mins = (TimeSpan)(attendanceRecord.TimeOut - attendanceRecord.TimeIn);

        //        if (attendanceRecord.StatusGZ == true)
        //        {
        //            attendanceRecord.GZOTMin = (short)mins.TotalMinutes;
        //            attendanceRecord.WorkMin = (short)mins.TotalMinutes;
        //            attendanceRecord.StatusGZOT = true;
        //        }
        //        else if (attendanceRecord.StatusDO == true)
        //        {
        //            attendanceRecord.OTMin = (short)mins.TotalMinutes;
        //            attendanceRecord.WorkMin = (short)mins.TotalMinutes;
        //            attendanceRecord.StatusOT = true;
        //        }
        //        else
        //        {
        //            //Calculate Late IN, Compare margin with Shift Late In
        //            if (attendanceRecord.TimeIn > roster.DutyTime)
        //            {
        //                TimeSpan lateMinsSpan = (TimeSpan)(attendanceRecord.TimeIn - roster.DutyTime);
        //                if (lateMinsSpan.TotalMinutes > context.Options.First().LateMin)
        //                {
        //                    attendanceRecord.LateIn = (short)lateMinsSpan.TotalMinutes;
        //                    attendanceRecord.StatusLI = true;
        //                    attendanceRecord.EarlyIn = null;
        //                }
        //            }

        //            //Calculate Early In, Compare margin with Shift Early In
        //            else if (attendanceRecord.TimeIn < roster.DutyTime)
        //            {
        //                TimeSpan EarlyInMinsSpan = (TimeSpan)(roster.DutyTime - attendanceRecord.TimeIn);
        //                if (EarlyInMinsSpan.TotalMinutes > context.Options.FirstOrDefault().EarlyIn)
        //                {
        //                    attendanceRecord.EarlyIn = (short)EarlyInMinsSpan.TotalMinutes;
        //                    attendanceRecord.StatusEI = true;
        //                    attendanceRecord.LateIn = null;
        //                }
        //            }

        //            // CalculateShiftEndTime = ShiftStart + DutyHours
        //            TimeSpan shiftEnd = roster.DutyTime + (new TimeSpan(0, (int)roster.WorkMin, 0));

        //            //Calculate Early Out, Compare margin with Shift Early Out
        //            if (attendanceRecord.TimeOut < shiftEnd)
        //            {
        //                TimeSpan EarlyOutMinsSpan = (TimeSpan)(shiftEnd - attendanceRecord.TimeOut);
        //                if (EarlyOutMinsSpan.TotalMinutes > context.Options.FirstOrDefault().EarlyOut)
        //                {
        //                    attendanceRecord.EarlyOut = (short)EarlyOutMinsSpan.TotalMinutes;
        //                    attendanceRecord.StatusEO = true;
        //                    attendanceRecord.LateOut = null;
        //                }
        //            }

        //            //Calculate Late Out, Compare margin with Shift Late Out
        //            if (attendanceRecord.TimeOut > shiftEnd)
        //            {
        //                TimeSpan LateOutMinsSpan = (TimeSpan)(attendanceRecord.TimeOut - shiftEnd);
        //                if (LateOutMinsSpan.TotalMinutes > context.Options.FirstOrDefault().LateMin)
        //                {
        //                    attendanceRecord.LateOut = (short)LateOutMinsSpan.TotalMinutes;
        //                    // Late Out cannot have an early out, In case of poll at multiple times before and after shiftend
        //                    attendanceRecord.EarlyOut = null;
        //                    attendanceRecord.StatusLO = true;
        //                }
        //            }

        //            if (attendanceRecord.EarlyIn != null && attendanceRecord.EarlyIn > context.Options.FirstOrDefault().EarlyIn)
        //            {
        //                attendanceRecord.WorkMin = (short)(mins.TotalMinutes - attendanceRecord.EarlyIn);
        //            }
        //            else if (attendanceRecord.LateOut != null && attendanceRecord.LateOut > context.Options.FirstOrDefault().OverTime)
        //            {
        //                attendanceRecord.WorkMin = (short)(mins.TotalMinutes - attendanceRecord.LateOut);
        //            }
        //            else
        //            {
        //                attendanceRecord.WorkMin = (short)(mins.TotalMinutes);
        //            }
        //            //Calculate OverTime, Compare margin with Shift OverTime
        //            if (attendanceRecord.EarlyIn > context.Options.FirstOrDefault().EarlyOut || attendanceRecord.LateOut > context.Options.FirstOrDefault().OverTime)
        //            {
        //                if (attendanceRecord.StatusGZ != true || attendanceRecord.StatusDO != true)
        //                {
        //                    short _EarlyIn;
        //                    short _LateOut;
        //                    if (attendanceRecord.EarlyIn == null)
        //                        _EarlyIn = 0;
        //                    else
        //                        _EarlyIn = (short)attendanceRecord.EarlyIn;

        //                    if (attendanceRecord.LateOut == null)
        //                        _LateOut = 0;
        //                    else
        //                        _LateOut = (short)attendanceRecord.LateOut;

        //                    attendanceRecord.OTMin = (short)(_EarlyIn + _LateOut);
        //                    attendanceRecord.StatusOT = true;
        //                }
        //            }
        //        }


        //    }
        //    catch (Exception ex)
        //    {

        //    }
        //}

        //private void CreateAttendance(DateTime dateTime)
        //{
        //    using (var ctx = new TAS2013Entities())
        //    {
        //        foreach (var emp in ctx.Emps.Where(aa => aa.ProcessAtt == true && aa.Status == true))
        //        {
        //            try
        //            {
        //                toolStripStatusLabel1.Text = "Creating Attendance For Date: " + dateTime.Date.ToString() + " EmpNo: " + emp.EmpNo;
        //                statusStrip1.Refresh();
        //                /////////////////////////////////////////////////////
        //                //  Mark Everyone Absent while creating Attendance //
        //                /////////////////////////////////////////////////////
        //                //Set DUTYCODE = D, StatusAB = true, and Remarks = [Absent]
        //                DataConvertor.NewDatabase.AttData att = new DataConvertor.NewDatabase.AttData();
        //                att.AttDate = dateTime.Date;
        //                att.DutyCode = "D";
        //                att.StatusAB = true;
        //                att.Remarks = "[Absent]";
        //                att.DutyTime = emp.Shift.StartTime;
        //                att.EmpID = emp.EmpID;
        //                att.EmpNo = emp.EmpNo;
        //                att.EmpDate = emp.EmpID + dateTime.ToString("yyMMdd");
        //                att.ShifMin = CalculateShiftMinutes(emp.Shift, dateTime.DayOfWeek);
        //                //////////////////////////
        //                //  Check for Rest Day //
        //                ////////////////////////
        //                //Set DutyCode = R, StatusAB=false, StatusDO = true, and Remarks=[DO]
        //                //Check for 1st Day Off of Shift
        //                if (emp.Shift.DaysName.Name == ReturnDayOfWeek(dateTime.DayOfWeek))
        //                {
        //                    att.DutyCode = "R";
        //                    att.StatusAB = false;
        //                    att.StatusDO = true;
        //                    att.Remarks = "[DO]";
        //                }
        //                //Check for 2nd Day Off of shift
        //                if (emp.Shift.DaysName1.Name == ReturnDayOfWeek(dateTime.DayOfWeek))
        //                {
        //                    att.DutyCode = "R";
        //                    att.StatusAB = false;
        //                    att.StatusDO = true;
        //                    att.Remarks = "[DO]";
        //                }

        //                //////////////////////////
        //                //  Check for GZ Day //
        //                ////////////////////////
        //                //Set DutyCode = R, StatusAB=false, StatusGZ = true, and Remarks=[GZ]
        //                if (emp.Shift.GZDays == true)
        //                {
        //                    foreach (var holiday in context.Holidays)
        //                    {
        //                        if (context.Holidays.Where(hol => hol.HolDate.Month == att.AttDate.Value.Month && hol.HolDate.Day == att.AttDate.Value.Day).Count() > 0)
        //                        {
        //                            att.DutyCode = "G";
        //                            att.StatusAB = false;
        //                            att.StatusGZ = true;
        //                            att.Remarks = "[GZ]";
        //                            att.ShifMin = 0;
        //                        }
        //                    }
        //                }

        //                //////////////////////////
        //                //  Check for Roster   //
        //                ////////////////////////
        //                //If Roster DutyCode is Rest then change the StatusAB and StatusDO
        //                foreach (var roster in context.Rosters.Where(aa => aa.EmpDate == att.EmpDate))
        //                {
        //                    att.DutyCode = roster.DutyCode;
        //                    if (roster.DutyCode == "R")
        //                    {
        //                        att.StatusAB = false;
        //                        att.StatusDO = true;
        //                        att.Remarks = "[DO]";
        //                    }

        //                    att.ShifMin = roster.WorkMin;
        //                    att.DutyTime = roster.DutyTime;
        //                }

        //                ////////////////////////////
        //                //TODO Check for Official Duty//
        //                //////////////////////////



        //                ////////////////////////////
        //                //  Check for Short Leave//
        //                //////////////////////////
        //                foreach (var sLeave in context.LvShorts)
        //                {
        //                    if (context.LvShorts.Where(lv => lv.EmpDate == att.EmpDate).Count() > 0)
        //                    {
        //                        att.StatusSL = true;
        //                        att.StatusAB = null;
        //                        att.DutyCode = "D";
        //                        att.Remarks = "[SL]";
        //                    }
        //                }

        //                //////////////////////////
        //                //   Check for Leave   //
        //                ////////////////////////
        //                //Set DutyCode = R, StatusAB=false, StatusGZ = true, and Remarks=[GZ]
        //                foreach (var Leave in context.LvDatas)
        //                {
        //                    if (context.LvDatas.Where(lv => lv.EmpDate == att.EmpDate).Count() > 0)
        //                    {
        //                        att.StatusLeave = true;
        //                        att.StatusAB = false;
        //                        att.DutyCode = "L";
        //                        att.Remarks = "[Leave]";
        //                    }
        //                    else
        //                    {
        //                        att.StatusLeave = false;
        //                    }
        //                }
        //                ctx.AttDatas.AddObject(att);

        //            }
        //            catch (Exception ex)
        //            {
        //                toolStripStatusLabel1.Text = "Exception at Creting Attendance, Date: " + dateTime.Date.ToString() + " EmpNo: " + emp.EmpNo;
        //                statusStrip1.Refresh();
        //                ListViewItem lvi = new ListViewItem(emp.EmpNo);
        //                listView1.Items.Add(lvi);
        //                listView1.Refresh();
        //            }
        //        }
        //        AttProcess attp = new AttProcess();
        //        attp.ProcessDate = dateTime;
        //        ctx.AttProcesses.AddObject(attp);
        //        ctx.SaveChanges();
        //    }


        //    toolStripStatusLabel1.Text = "Creat Attendance Process Complete: ";
        //    statusStrip1.Refresh();
        //}

        //private string ReturnDayOfWeek(DayOfWeek dayOfWeek)
        //{
        //    string _DayName = "";
        //    switch (dayOfWeek)
        //    {
        //        case DayOfWeek.Monday:
        //            _DayName = "Monday";
        //            break;
        //        case DayOfWeek.Tuesday:
        //            _DayName = "Tuesday";
        //            break;
        //        case DayOfWeek.Wednesday:
        //            _DayName = "Wednesday";
        //            break;
        //        case DayOfWeek.Thursday:
        //            _DayName = "Thursday";
        //            break;
        //        case DayOfWeek.Friday:
        //            _DayName = "Friday";
        //            break;
        //        case DayOfWeek.Saturday:
        //            _DayName = "Saturday";
        //            break;
        //        case DayOfWeek.Sunday:
        //            _DayName = "Sunday";
        //            break;
        //    }
        //    return _DayName;
        //}

        //private TimeSpan CalculateShiftEndTime(Shift shift, DayOfWeek dayOfWeek)
        //{
        //    Int16 workMins = 0;
        //    try
        //    {
        //        switch (dayOfWeek)
        //        {
        //            case DayOfWeek.Monday:
        //                workMins = shift.MonMin;
        //                break;
        //            case DayOfWeek.Tuesday:
        //                workMins = shift.TueMin;
        //                break;
        //            case DayOfWeek.Wednesday:
        //                workMins = shift.WedMin;
        //                break;
        //            case DayOfWeek.Thursday:
        //                workMins = shift.ThuMin;
        //                break;
        //            case DayOfWeek.Friday:
        //                workMins = shift.FriMin;
        //                break;
        //            case DayOfWeek.Saturday:
        //                workMins = shift.SatMin;
        //                break;
        //            case DayOfWeek.Sunday:
        //                workMins = shift.SunMin;
        //                break;
        //        }
        //    }
        //    catch (Exception ex)
        //    {

        //    }
        //    return shift.StartTime + (new TimeSpan(0, workMins, 0));
        //}

        //private Int16 CalculateShiftMinutes(Shift shift, DayOfWeek dayOfWeek)
        //{
        //    Int16 workMins = 0;
        //    try
        //    {
        //        switch (dayOfWeek)
        //        {
        //            case DayOfWeek.Monday:
        //                workMins = shift.MonMin;
        //                break;
        //            case DayOfWeek.Tuesday:
        //                workMins = shift.TueMin;
        //                break;
        //            case DayOfWeek.Wednesday:
        //                workMins = shift.WedMin;
        //                break;
        //            case DayOfWeek.Thursday:
        //                workMins = shift.ThuMin;
        //                break;
        //            case DayOfWeek.Friday:
        //                workMins = shift.FriMin;
        //                break;
        //            case DayOfWeek.Saturday:
        //                workMins = shift.SatMin;
        //                break;
        //            case DayOfWeek.Sunday:
        //                workMins = shift.SunMin;
        //                break;
        //        }
        //    }
        //    catch (Exception ex)
        //    {

        //    }
        //    return workMins;
        //}
    }
}
