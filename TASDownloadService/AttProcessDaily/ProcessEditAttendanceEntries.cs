using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using TASDownloadService.Model;

namespace WMSFFService
{
    public class ProcessEditAttendanceEntries
    {
        TAS2013Entities newDB = new TAS2013Entities();
        MyCustomFunctions _myHelperClass = new MyCustomFunctions();
        public void ProcessManualEditAttendance(DateTime _dateStart, DateTime _dateEnd)
        {
            List<AttDataManEdit> _attEdit = new List<AttDataManEdit>();
            List<AttData> _AttData = new List<AttData>();
            AttData _TempAttData = new AttData();
            using (var ctx = new TAS2013Entities())
            {

                if (_dateStart == _dateEnd)
                {
                    _attEdit = ctx.AttDataManEdits.Where(aa => aa.NewTimeIn == _dateStart).OrderBy(aa => aa.EditDateTime).ToList();
                    _dateEnd = _dateEnd + new TimeSpan(23, 59, 59);
                    //_attEdit = ctx.AttDataManEdits.Where(aa => aa.NewTimeIn >= _dateStart && aa.NewTimeIn <= _dateEnd && (aa.EmpID == 472)).OrderBy(aa => aa.EditDateTime).ToList();
                    _AttData = ctx.AttDatas.Where(aa => aa.AttDate == _dateStart).ToList();
                }
                else
                {
                    _attEdit = ctx.AttDataManEdits.Where(aa => aa.NewTimeIn >= _dateStart && aa.NewTimeOut <= _dateEnd).OrderBy(aa => aa.EditDateTime).ToList();
                    //_attEdit = ctx.AttDataManEdits.Where(aa => aa.NewTimeIn >= _dateStart && (aa.NewTimeOut <= _dateEnd && aa.EmpID == 472)).OrderBy(aa => aa.EditDateTime).ToList();
                    _AttData = ctx.AttDatas.Where(aa => aa.AttDate >= _dateStart && aa.AttDate <= _dateEnd).ToList();
                }


                foreach (var item in _attEdit)
                {
                    _TempAttData = _AttData.First(aa => aa.EmpDate == item.EmpDate);
                    _TempAttData.TimeIn = item.NewTimeIn;
                    _TempAttData.TimeOut = item.NewTimeOut;
                    _TempAttData.DutyCode = item.NewDutyCode;
                    _TempAttData.DutyTime = item.NewDutyTime;
                    switch (_TempAttData.DutyCode)
                    {
                        case "D":
                            _TempAttData.StatusAB = true;
                            _TempAttData.StatusP = false;
                            _TempAttData.StatusMN = true;
                            _TempAttData.StatusDO = false;
                            _TempAttData.StatusGZ = false;
                            _TempAttData.StatusLeave = false;
                            _TempAttData.StatusOT = false;
                            _TempAttData.OTMin = null;
                            _TempAttData.EarlyIn = null;
                            _TempAttData.EarlyOut = null;
                            _TempAttData.LateIn = null;
                            _TempAttData.LateOut = null;
                            _TempAttData.WorkMin = null;
                            _TempAttData.GZOTMin = null;
                            break;
                        case "G":
                            _TempAttData.StatusAB = false;
                            _TempAttData.StatusP = false;
                            _TempAttData.StatusMN = true;
                            _TempAttData.StatusDO = false;
                            _TempAttData.StatusGZ = true;
                            _TempAttData.StatusLeave = false;
                            _TempAttData.StatusOT = false;
                            _TempAttData.OTMin = null;
                            _TempAttData.EarlyIn = null;
                            _TempAttData.EarlyOut = null;
                            _TempAttData.LateIn = null;
                            _TempAttData.LateOut = null;
                            _TempAttData.WorkMin = null;
                            _TempAttData.GZOTMin = null;
                            break;
                        case "R":
                            _TempAttData.StatusAB = false;
                            _TempAttData.StatusP = false;
                            _TempAttData.StatusMN = true;
                            _TempAttData.StatusDO = true;
                            _TempAttData.StatusGZ = false;
                            _TempAttData.StatusLeave = false;
                            _TempAttData.StatusOT = false;
                            _TempAttData.OTMin = null;
                            _TempAttData.EarlyIn = null;
                            _TempAttData.EarlyOut = null;
                            _TempAttData.LateIn = null;
                            _TempAttData.LateOut = null;
                            _TempAttData.WorkMin = null;
                            _TempAttData.GZOTMin = null;
                            break;
                    }
                    if (_TempAttData.TimeIn != null && _TempAttData.TimeOut != null)
                    {
                        //If TimeIn = TimeOut then calculate according to DutyCode
                        if (_TempAttData.TimeIn == _TempAttData.TimeOut)
                        {
                            CalculateInEqualToOut(_TempAttData);
                        }
                        else
                        {
                            if (_TempAttData.DutyTime == new TimeSpan(0, 0, 0))
                            {
                                CalculateOpenShiftTimes(_TempAttData, _TempAttData.Emp.Shift);
                            }
                            else
                            {
                                //if (attendanceRecord.TimeIn.Value.Date.Day == attendanceRecord.TimeOut.Value.Date.Day)
                                //{
                                CalculateShiftTimes(_TempAttData, _TempAttData.Emp.Shift);
                                //}
                                //else
                                //{
                                //    CalculateOpenShiftTimes(attendanceRecord, shift);
                                //}
                            }
                        }

                        //If TimeIn = TimeOut then calculate according to DutyCode
                    }
                    ctx.SaveChanges();
                }
                ctx.Dispose();


            }
            _myHelperClass.WriteToLogFile("ProcessManual Attendance Completed: ");
        }
        private void CalculateInEqualToOut(AttData attendanceRecord)
        {
            switch (attendanceRecord.DutyCode)
            {
                case "G":
                    attendanceRecord.StatusAB = false;
                    attendanceRecord.StatusGZ = true;
                    attendanceRecord.WorkMin = 0;
                    attendanceRecord.EarlyIn = 0;
                    attendanceRecord.EarlyOut = 0;
                    attendanceRecord.LateIn = 0;
                    attendanceRecord.LateOut = 0;
                    attendanceRecord.OTMin = 0;
                    attendanceRecord.GZOTMin = 0;
                    attendanceRecord.StatusGZOT = false;
                    attendanceRecord.TimeIn = null;
                    attendanceRecord.TimeOut = null;
                    attendanceRecord.Remarks = "[GZ][Manual]";
                    break;
                case "R":
                    attendanceRecord.StatusAB = false;
                    attendanceRecord.StatusGZ = false;
                    attendanceRecord.WorkMin = 0;
                    attendanceRecord.EarlyIn = 0;
                    attendanceRecord.EarlyOut = 0;
                    attendanceRecord.LateIn = 0;
                    attendanceRecord.LateOut = 0;
                    attendanceRecord.OTMin = 0;
                    attendanceRecord.GZOTMin = 0;
                    attendanceRecord.StatusGZOT = false;
                    attendanceRecord.TimeIn = null;
                    attendanceRecord.TimeOut = null;
                    attendanceRecord.StatusDO = true;
                    attendanceRecord.Remarks = "[DO][Manual]";
                    break;
                case "D":
                    attendanceRecord.StatusAB = true;
                    attendanceRecord.StatusGZ = false;
                    attendanceRecord.WorkMin = 0;
                    attendanceRecord.EarlyIn = 0;
                    attendanceRecord.EarlyOut = 0;
                    attendanceRecord.LateIn = 0;
                    attendanceRecord.LateOut = 0;
                    attendanceRecord.OTMin = 0;
                    attendanceRecord.GZOTMin = 0;
                    attendanceRecord.StatusGZOT = false;
                    attendanceRecord.TimeIn = null;
                    attendanceRecord.TimeOut = null;
                    attendanceRecord.StatusDO = false;
                    attendanceRecord.StatusP = false;
                    attendanceRecord.Remarks = "[Absent][Manual]";
                    break;
            }
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
                string _error = "";
                if (ex.InnerException.Message != null)
                    _error = ex.InnerException.Message;
                else
                    _error = ex.Message;
                _myHelperClass.WriteToLogFile("Manual Attendance Processing 4" + _error);
            }
            return workMins;
        }
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

                        //Subtract EarlyIn and LateOut from Work Minutes
                        //////-------to-do--------- Automate earlyin,lateout from shift setup
                        attendanceRecord.WorkMin = (short)(mins.TotalMinutes);
                        if (attendanceRecord.EarlyIn != null && attendanceRecord.EarlyIn > shift.EarlyIn)
                        {
                            attendanceRecord.WorkMin = (short)(attendanceRecord.WorkMin - attendanceRecord.EarlyIn);
                        }
                        if (attendanceRecord.LateOut != null && attendanceRecord.LateOut > shift.LateOut)
                        {
                            attendanceRecord.WorkMin = (short)(attendanceRecord.WorkMin - attendanceRecord.LateOut);
                        }
                        if (attendanceRecord.LateOut != null || attendanceRecord.EarlyIn != null)

                            // round off work mins if overtime less than shift.OverTimeMin >
                            if (attendanceRecord.WorkMin > CalculateShiftMinutes(shift, attendanceRecord.AttDate.Value.DayOfWeek) && (attendanceRecord.WorkMin <= (CalculateShiftMinutes(shift, attendanceRecord.AttDate.Value.DayOfWeek) + shift.OverTimeMin)))
                            {
                                attendanceRecord.WorkMin = CalculateShiftMinutes(shift, attendanceRecord.AttDate.Value.DayOfWeek);
                            }
                        //Calculate OverTime = OT, Compare margin with Shift OverTime
                        //----to-do----- Handle from shift
                        //if (attendanceRecord.EarlyIn > shift.EarlyIn || attendanceRecord.LateOut > shift.LateOut)
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
                        //        attendanceRecord.Remarks = attendanceRecord.Remarks + "[N-OT]";
                        //    }
                        //}
                        if (attendanceRecord.StatusGZ != true || attendanceRecord.StatusDO != true)
                        {
                            if (attendanceRecord.LateOut != null)
                            {
                                attendanceRecord.OTMin = attendanceRecord.LateOut;
                                attendanceRecord.StatusOT = true;
                                attendanceRecord.Remarks = attendanceRecord.Remarks + "[N-OT]";
                            }
                            else
                            {

                                attendanceRecord.OTMin = null;
                                attendanceRecord.StatusOT = null;
                                attendanceRecord.Remarks.Replace("[OT]", "");
                                attendanceRecord.Remarks.Replace("[N-OT]", "");
                            }
                        }
                        else
                        {
                            attendanceRecord.OTMin = null;
                            attendanceRecord.StatusOT = null;
                            attendanceRecord.Remarks.Replace("[OT]", "");
                            attendanceRecord.Remarks.Replace("[N-OT]", "");
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
                                // CalculateShiftEndTime = ShiftStart + DutyHours
                                TimeSpan shiftEnd = CalculateShiftEndTime(shift, attendanceRecord.AttDate.Value.DayOfWeek);
                                attendanceRecord.WorkMin = (short)(mins.TotalMinutes);
                                //Calculate OverTIme, 
                                if (mins.TotalMinutes > (CalculateShiftMinutes(shift, attendanceRecord.AttDate.Value.DayOfWeek) + shift.OverTimeMin))
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
        }
        #endregion


    }
}
