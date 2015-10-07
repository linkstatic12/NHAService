using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using TASDownloadService.Model;

namespace WMSFFService
{
    //public class ProcessManualAttendance
    //{
    //    TAS_FF2000Entities oldDB = new TAS_FF2000Entities();
    //    TAS2013Entities newDB = new TAS2013Entities();
    //    MyCustomFunctions _myHelperClass = new MyCustomFunctions();
    //    public void ProcessManualEditAttendance()
    //    {
    //        using (var ctx = new TAS2013Entities())
    //        {
    //            List<NewDatabase.AttData> _NewdbAttData = new List<NewDatabase.AttData>();
    //            List<OldDatabase.OldModel.AttData> _OldDBAttData = new List<OldDatabase.OldModel.AttData>();
    //            DateTime _dt = GlobalSettings._dateTime;
    //            oldDB.CommandTimeout = 0;
    //            _OldDBAttData = oldDB.AttDatas.Where(aa => aa.AttDate >= _dt.Date && aa.StatusMN == 1).ToList();
    //            _NewdbAttData = ctx.AttDatas.Where(aa => aa.AttDate >= _dt).ToList();
    //            AttDataManEdit _AttDataManEdit = new AttDataManEdit();
    //            _myHelperClass.WriteToLogFile("ProcessManual Attendance Function: ");
    //            foreach (var _attendanceRecord in _OldDBAttData)
    //            {
    //                try
    //                {
    //                    long _empDate = Convert.ToInt64(_attendanceRecord.EmpDate);
    //                    NewDatabase.AttData editedAttData = _NewdbAttData.FirstOrDefault(aa => aa.EmpDate == _empDate.ToString());
    //                    if (editedAttData != null)
    //                    {
    //                        if (_attendanceRecord.Time0 != "" && _attendanceRecord.Time1 != "")
    //                        {
    //                            // Put Old Entry To AttEdit OldTime
    //                            //Becuase we are getting R or 075 values from old data
    //                            bool check = false;
    //                            if (_attendanceRecord.Time0.Length == 4)
    //                            {
    //                                _AttDataManEdit = SaveOldAttTimesInOut(editedAttData);
    //                                editedAttData.TimeIn = _attendanceRecord.AttDate + GlobalSettings.ConvertTime(_attendanceRecord.Time0);
    //                            }
    //                            TimeSpan OpenShiftThresholdStart = new TimeSpan(17, 00, 00);

    //                            TimeSpan OpenShiftThresholdEnd = new TimeSpan(11, 00, 00);
    //                            if (editedAttData.TimeIn.Value.TimeOfDay > OpenShiftThresholdStart)
    //                            {
    //                                DateTime dt = new DateTime();
    //                                dt = _attendanceRecord.AttDate.Value.Date.AddDays(1);
    //                                if (_attendanceRecord.Time1.Length == 4)
    //                                    editedAttData.TimeOut = dt + GlobalSettings.ConvertTime(_attendanceRecord.Time1);
    //                                check = true;
    //                            }
    //                            else
    //                            {
    //                                if (_attendanceRecord.Time1.Length == 4)
    //                                        editedAttData.TimeOut = _attendanceRecord.AttDate.Value + GlobalSettings.ConvertTime(_attendanceRecord.Time1);
    //                            }
    //                            if (_attendanceRecord.Remarks != null)
    //                                editedAttData.Remarks = _attendanceRecord.Remarks + "[Manual]";
    //                            try
    //                            {
    //                                SaveNewAttTimeInOut(_attendanceRecord, _AttDataManEdit);
    //                                ctx.AttDataManEdits.AddObject(_AttDataManEdit);
    //                                ChangeManualAttData(_AttDataManEdit, editedAttData, check);
    //                            }
    //                            catch (Exception ex)
    //                            {

    //                            }
                                
    //                        }
    //                        else
    //                        {
    //                            switch (_attendanceRecord.Remarks)
    //                            {
    //                                case "[Manual][D/Off]":
    //                                    editedAttData.StatusAB = false;
    //                                    editedAttData.StatusDO = true;
    //                                    editedAttData.Remarks = "[DO]";
    //                                    editedAttData.DutyCode = "R";
    //                                    editedAttData.StatusMN = true;
    //                                    break;
    //                                case "[Day Off][Manual]":
    //                                    editedAttData.StatusAB = false;
    //                                    editedAttData.StatusDO = true;
    //                                    editedAttData.Remarks = "[DO]";
    //                                    editedAttData.DutyCode = "R";
    //                                    editedAttData.StatusMN = true;
    //                                    break;
    //                                case "[DO]":
    //                                    editedAttData.StatusAB = false;
    //                                    editedAttData.StatusDO = true;
    //                                    editedAttData.Remarks = "[DO]";
    //                                    editedAttData.DutyCode = "R";
    //                                    editedAttData.StatusMN = true;
    //                                    break;
    //                                case "[Manual][O/Hol]":
    //                                    editedAttData.StatusAB = false;
    //                                    editedAttData.StatusGZ = true;
    //                                    editedAttData.Remarks = "[GZ]";
    //                                    editedAttData.DutyCode = "G";
    //                                    editedAttData.StatusMN = true;
    //                                    break;
    //                                case "[Leave][Manual]":
    //                                    editedAttData.StatusAB = false;
    //                                    editedAttData.StatusLeave = true;
    //                                    editedAttData.Remarks = "[Leave]";
    //                                    editedAttData.DutyCode = "D";
    //                                    editedAttData.StatusMN = true;
    //                                    break;
    //                                case "[Absent]":
    //                                    editedAttData.StatusAB = true;
    //                                    editedAttData.Remarks = "[Absent]";
    //                                    editedAttData.DutyCode = "D";
    //                                    editedAttData.StatusMN = true;
    //                                    break;
    //                            }
    //                        }
    //                    }
    //                    ctx.SaveChanges();
    //                }
    //                catch (Exception ex)
    //                {
    //                    string _error = "";
    //                    if (ex.InnerException.Message != null)
    //                        _error = ex.InnerException.Message;
    //                    else
    //                        _error = ex.Message;
    //                    _myHelperClass.WriteToLogFile("Manual Attendance Processing 1" + _error);
    //                }
    //            }
    //        }
    //        _myHelperClass.WriteToLogFile("ProcessManual Attendance Completed: ");
    //    }

    //    private void ChangeManualAttData(AttDataManEdit _AttDataManEdit, NewDatabase.AttData editedAttData,bool _check)
    //    {
    //        //true for open shift
    //        if (_check == false)
    //        {
    //            switch (_AttDataManEdit.NewDutyCode)
    //            {
    //                case "D":
    //                    editedAttData.DutyCode = "D";
    //                    if (_AttDataManEdit.NewTimeIn != null && _AttDataManEdit.NewTimeOut != null)
    //                    {
    //                        CalculateShiftTimes(editedAttData, editedAttData.Emp.Shift);
    //                        editedAttData.StatusAB = false;
    //                        editedAttData.StatusP = true;
    //                        editedAttData.StatusMN = true;
    //                    }
    //                    break;
    //                case "G":
    //                    editedAttData.DutyCode = "G";
    //                    editedAttData.StatusGZ = true;
    //                    if (_AttDataManEdit.NewTimeIn != null && _AttDataManEdit.NewTimeOut != null)
    //                    {
    //                        CalculateShiftTimes(editedAttData, editedAttData.Emp.Shift);
    //                        editedAttData.StatusP = true;
    //                    }
    //                    editedAttData.StatusAB = false;
    //                    editedAttData.StatusMN = true;
    //                    break;
    //                case "R":
    //                    editedAttData.DutyCode = "R";
    //                    editedAttData.StatusDO = true;
    //                    if (_AttDataManEdit.NewTimeIn != null && _AttDataManEdit.NewTimeOut != null)
    //                    {
    //                        CalculateShiftTimes(editedAttData, editedAttData.Emp.Shift);
    //                        editedAttData.StatusP = true;
    //                    }
    //                    editedAttData.StatusAB = false;
    //                    editedAttData.StatusMN = true;
    //                    break;
    //            }
    //        }
    //        else
    //        {
    //            switch (_AttDataManEdit.NewDutyCode)
    //            {
    //                case "D":
    //                    editedAttData.DutyCode = "D";
    //                    if (_AttDataManEdit.NewTimeIn != null && _AttDataManEdit.NewTimeOut != null)
    //                    {
    //                        CalculateOpenShiftTimes(editedAttData, editedAttData.Emp.Shift);
    //                        editedAttData.StatusAB = false;
    //                        editedAttData.StatusP = true;
    //                        editedAttData.StatusMN = true;
    //                    }
    //                    break;
    //                case "G":
    //                    editedAttData.DutyCode = "G";
    //                    editedAttData.StatusGZ = true;
    //                    if (_AttDataManEdit.NewTimeIn != null && _AttDataManEdit.NewTimeOut != null)
    //                    {
    //                        CalculateOpenShiftTimes(editedAttData, editedAttData.Emp.Shift);
    //                        editedAttData.StatusP = true;
    //                    }
    //                    editedAttData.StatusAB = false;
    //                    editedAttData.StatusMN = true;
    //                    break;
    //                case "R":
    //                    editedAttData.DutyCode = "R";
    //                    editedAttData.StatusDO = true;
    //                    if (_AttDataManEdit.NewTimeIn != null && _AttDataManEdit.NewTimeOut != null)
    //                    {
    //                        CalculateOpenShiftTimes(editedAttData, editedAttData.Emp.Shift);
    //                        editedAttData.StatusP = true;
    //                    }
    //                    editedAttData.StatusAB = false;
    //                    editedAttData.StatusMN = true;
    //                    break;
    //            }
    //        }
            
    //    }

    //    private void SaveNewAttTimeInOut(OldDatabase.OldModel.AttData _attendanceRecord, AttDataManEdit _AttDataManEdit)
    //    {
    //        try
    //        {
    //            if (_attendanceRecord.Time0.Length == 4)
    //                _AttDataManEdit.NewTimeIn = _attendanceRecord.AttDate + GlobalSettings.ConvertTime(_attendanceRecord.Time0);
    //            if (_attendanceRecord.Time1.Length == 4)
    //                _AttDataManEdit.NewTimeOut = _attendanceRecord.AttDate + GlobalSettings.ConvertTime(_attendanceRecord.Time1);
    //            if (_attendanceRecord.DutyTime.Length == 4)
    //                _AttDataManEdit.NewDutyTime = GlobalSettings.ConvertTime(_attendanceRecord.DutyTime);
    //            if (_attendanceRecord.DutyCode == "DUTY")
    //                _AttDataManEdit.NewDutyCode = "D";
    //            else if (_attendanceRecord.DutyCode == "G")
    //                _AttDataManEdit.NewDutyCode = "G";
    //            else if (_attendanceRecord.DutyCode == "R")
    //                _AttDataManEdit.NewDutyCode = "R";
    //            _AttDataManEdit.NewRemarks = _attendanceRecord.Remarks;
    //        }
    //        catch (Exception ex)
    //        {
    //            string _error = "";
    //            if (ex.InnerException.Message != null)
    //                _error = ex.InnerException.Message;
    //            else
    //                _error = ex.Message;
    //            _myHelperClass.WriteToLogFile("Manual Attendance Processing 2" + _error);
    //        }
    //    }

    //    private AttDataManEdit SaveOldAttTimesInOut(AttData _NewdbAttData)
    //    {
    //        AttDataManEdit _AttEdit = new AttDataManEdit();
    //        _AttEdit.EmpID = _NewdbAttData.EmpID;
    //        _AttEdit.EmpDate = _NewdbAttData.EmpDate;
    //        if (_NewdbAttData.TimeIn != null)
    //            _AttEdit.OldTimeIn = _NewdbAttData.TimeIn;
    //        if(_NewdbAttData.TimeOut!=null)
    //            _AttEdit.OldTimeOut = _NewdbAttData.TimeOut;
    //        _AttEdit.OldDutyCode = _NewdbAttData.DutyCode;
    //        _AttEdit.OldDutyTime = _NewdbAttData.DutyTime;
    //        _AttEdit.EditDateTime = DateTime.Now;
    //        _AttEdit.OldRemarks = _NewdbAttData.Remarks;
    //        _AttEdit.OldDutyTime = _NewdbAttData.DutyTime;
    //        return _AttEdit;
    //    }

    //    private TimeSpan CalculateShiftEndTime(Shift shift, DayOfWeek dayOfWeek)
    //    {
    //        Int16 workMins = 0;
    //        try
    //        {
    //            switch (dayOfWeek)
    //            {
    //                case DayOfWeek.Monday:
    //                    workMins = shift.MonMin;
    //                    break;
    //                case DayOfWeek.Tuesday:
    //                    workMins = shift.TueMin;
    //                    break;
    //                case DayOfWeek.Wednesday:
    //                    workMins = shift.WedMin;
    //                    break;
    //                case DayOfWeek.Thursday:
    //                    workMins = shift.ThuMin;
    //                    break;
    //                case DayOfWeek.Friday:
    //                    workMins = shift.FriMin;
    //                    break;
    //                case DayOfWeek.Saturday:
    //                    workMins = shift.SatMin;
    //                    break;
    //                case DayOfWeek.Sunday:
    //                    workMins = shift.SunMin;
    //                    break;
    //            }
    //        }
    //        catch (Exception ex)
    //        {

    //        }
    //        return shift.StartTime + (new TimeSpan(0, workMins, 0));
    //    }

    //    private Int16 CalculateShiftMinutes(Shift shift, DayOfWeek dayOfWeek)
    //    {
    //        Int16 workMins = 0;
    //        try
    //        {
    //            switch (dayOfWeek)
    //            {
    //                case DayOfWeek.Monday:
    //                    workMins = shift.MonMin;
    //                    break;
    //                case DayOfWeek.Tuesday:
    //                    workMins = shift.TueMin;
    //                    break;
    //                case DayOfWeek.Wednesday:
    //                    workMins = shift.WedMin;
    //                    break;
    //                case DayOfWeek.Thursday:
    //                    workMins = shift.ThuMin;
    //                    break;
    //                case DayOfWeek.Friday:
    //                    workMins = shift.FriMin;
    //                    break;
    //                case DayOfWeek.Saturday:
    //                    workMins = shift.SatMin;
    //                    break;
    //                case DayOfWeek.Sunday:
    //                    workMins = shift.SunMin;
    //                    break;
    //            }
    //        }
    //        catch (Exception ex)
    //        {
    //            string _error = "";
    //            if (ex.InnerException.Message != null)
    //                _error = ex.InnerException.Message;
    //            else
    //                _error = ex.Message;
    //            _myHelperClass.WriteToLogFile("Manual Attendance Processing 4" + _error);
    //        }
    //        return workMins;
    //    }

    //    private void CalculateShiftTimes(AttData attendanceRecord, Shift shift)
    //    {
    //        try
    //        {
    //            //Calculate WorkMin
    //            attendanceRecord.Remarks = "";
    //            TimeSpan mins = (TimeSpan)(attendanceRecord.TimeOut - attendanceRecord.TimeIn);
    //            //Check if GZ holiday then place all WorkMin in GZOTMin
    //            if (attendanceRecord.StatusGZ == true)
    //            {
    //                attendanceRecord.GZOTMin = (short)mins.TotalMinutes;
    //                attendanceRecord.WorkMin = (short)mins.TotalMinutes;
    //                attendanceRecord.StatusGZOT = true;
    //                attendanceRecord.Remarks = attendanceRecord.Remarks + "[G-OT]";
    //            }
    //            //if Rest day then place all WorkMin in OTMin
    //            else if (attendanceRecord.StatusDO == true)
    //            {
    //                attendanceRecord.OTMin = (short)mins.TotalMinutes;
    //                attendanceRecord.WorkMin = (short)mins.TotalMinutes;
    //                attendanceRecord.StatusOT = true;
    //                attendanceRecord.Remarks = attendanceRecord.Remarks + "[R-OT]";
    //            }
    //            else
    //            {
    //                /////////// to-do -----calculate Margins for those shifts which has break mins 
    //                if (shift.HasBreak == true)
    //                {
    //                    attendanceRecord.WorkMin = (short)(mins.TotalMinutes - shift.BreakMin);
    //                    attendanceRecord.ShifMin = (short)(CalculateShiftMinutes(shift, attendanceRecord.AttDate.Value.DayOfWeek) - (short)shift.BreakMin);
    //                }
    //                else
    //                {
    //                    //Calculate Late IN, Compare margin with Shift Late In
    //                    if (attendanceRecord.TimeIn.Value.TimeOfDay > shift.StartTime)
    //                    {
    //                        TimeSpan lateMinsSpan = (TimeSpan)(attendanceRecord.TimeIn.Value.TimeOfDay - shift.StartTime);
    //                        if (lateMinsSpan.TotalMinutes > shift.LateIn)
    //                        {
    //                            attendanceRecord.LateIn = (short)lateMinsSpan.TotalMinutes;
    //                            attendanceRecord.StatusLI = true;
    //                            attendanceRecord.EarlyIn = null;
    //                            attendanceRecord.Remarks = attendanceRecord.Remarks + "[LI]";
    //                        }
    //                    }
    //                    else
    //                    {
    //                        attendanceRecord.StatusLI = null;
    //                        attendanceRecord.LateIn = null;
    //                        attendanceRecord.Remarks.Replace("[LI]", "");
    //                    }

    //                    //Calculate Early In, Compare margin with Shift Early In
    //                    if (attendanceRecord.TimeIn.Value.TimeOfDay < shift.StartTime)
    //                    {
    //                        TimeSpan EarlyInMinsSpan = (TimeSpan)(shift.StartTime - attendanceRecord.TimeIn.Value.TimeOfDay);
    //                        if (EarlyInMinsSpan.TotalMinutes > shift.EarlyIn)
    //                        {
    //                            attendanceRecord.EarlyIn = (short)EarlyInMinsSpan.TotalMinutes;
    //                            attendanceRecord.StatusEI = true;
    //                            attendanceRecord.LateIn = null;
    //                            attendanceRecord.Remarks = attendanceRecord.Remarks + "[EI]";
    //                        }
    //                    }
    //                    else
    //                    {
    //                        attendanceRecord.StatusEI = null;
    //                        attendanceRecord.EarlyIn = null;
    //                        attendanceRecord.Remarks.Replace("[EI]", "");
    //                    }

    //                    // CalculateShiftEndTime = ShiftStart + DutyHours
    //                    TimeSpan shiftEnd = CalculateShiftEndTime(shift, attendanceRecord.AttDate.Value.DayOfWeek);

    //                    //Calculate Early Out, Compare margin with Shift Early Out
    //                    if (attendanceRecord.TimeOut.Value.TimeOfDay < shiftEnd)
    //                    {
    //                        TimeSpan EarlyOutMinsSpan = (TimeSpan)(shiftEnd - attendanceRecord.TimeOut.Value.TimeOfDay);
    //                        if (EarlyOutMinsSpan.TotalMinutes > shift.EarlyOut)
    //                        {
    //                            attendanceRecord.EarlyOut = (short)EarlyOutMinsSpan.TotalMinutes;
    //                            attendanceRecord.StatusEO = true;
    //                            attendanceRecord.LateOut = null;
    //                            attendanceRecord.Remarks = attendanceRecord.Remarks + "[EO]";
    //                        }
    //                    }
    //                    else
    //                    {
    //                        attendanceRecord.StatusEO = null;
    //                        attendanceRecord.EarlyOut = null;
    //                        attendanceRecord.Remarks.Replace("[EO]", "");
    //                    }
    //                    //Calculate Late Out, Compare margin with Shift Late Out
    //                    if (attendanceRecord.TimeOut.Value.TimeOfDay > shiftEnd)
    //                    {
    //                        TimeSpan LateOutMinsSpan = (TimeSpan)(attendanceRecord.TimeOut.Value.TimeOfDay - shiftEnd);
    //                        if (LateOutMinsSpan.TotalMinutes > shift.LateOut)
    //                        {
    //                            attendanceRecord.LateOut = (short)LateOutMinsSpan.TotalMinutes;
    //                            // Late Out cannot have an early out, In case of poll at multiple times before and after shiftend
    //                            attendanceRecord.EarlyOut = null;
    //                            attendanceRecord.StatusLO = true;
    //                            attendanceRecord.Remarks = attendanceRecord.Remarks + "[LO]";
    //                        }
    //                    }
    //                    else
    //                    {
    //                        attendanceRecord.StatusLO = null;
    //                        attendanceRecord.LateOut = null;
    //                        attendanceRecord.Remarks.Replace("[LO]", "");
                            
    //                    }

    //                    //Subtract EarlyIn and LateOut from Work Minutes
    //                    //////-------to-do--------- Automate earlyin,lateout from shift setup
    //                    attendanceRecord.WorkMin = (short)(mins.TotalMinutes);
    //                    if (attendanceRecord.EarlyIn != null && attendanceRecord.EarlyIn > shift.EarlyIn)
    //                    {
    //                        attendanceRecord.WorkMin = (short)(attendanceRecord.WorkMin - attendanceRecord.EarlyIn);
    //                    }
    //                    if (attendanceRecord.LateOut != null && attendanceRecord.LateOut > shift.LateOut)
    //                    {
    //                        attendanceRecord.WorkMin = (short)(attendanceRecord.WorkMin - attendanceRecord.LateOut);
    //                    }
    //                    if (attendanceRecord.LateOut != null || attendanceRecord.EarlyIn != null)

    //                        // round off work mins if overtime less than shift.OverTimeMin >
    //                        if (attendanceRecord.WorkMin > CalculateShiftMinutes(shift, attendanceRecord.AttDate.Value.DayOfWeek) && (attendanceRecord.WorkMin <= (CalculateShiftMinutes(shift, attendanceRecord.AttDate.Value.DayOfWeek) + shift.OverTimeMin)))
    //                        {
    //                            attendanceRecord.WorkMin = CalculateShiftMinutes(shift, attendanceRecord.AttDate.Value.DayOfWeek);
    //                        }
    //                    //Calculate OverTime = OT, Compare margin with Shift OverTime
    //                    //----to-do----- Handle from shift
    //                    //if (attendanceRecord.EarlyIn > shift.EarlyIn || attendanceRecord.LateOut > shift.LateOut)
    //                    //{
    //                    //    if (attendanceRecord.StatusGZ != true || attendanceRecord.StatusDO != true)
    //                    //    {
    //                    //        short _EarlyIn;
    //                    //        short _LateOut;
    //                    //        if (attendanceRecord.EarlyIn == null)
    //                    //            _EarlyIn = 0;
    //                    //        else
    //                    //            _EarlyIn = 0;

    //                    //        if (attendanceRecord.LateOut == null)
    //                    //            _LateOut = 0;
    //                    //        else
    //                    //            _LateOut = (short)attendanceRecord.LateOut;

    //                    //        attendanceRecord.OTMin = (short)(_EarlyIn + _LateOut);
    //                    //        attendanceRecord.StatusOT = true;
    //                    //        attendanceRecord.Remarks = attendanceRecord.Remarks + "[N-OT]";
    //                    //    }
    //                    //}
    //                    if (attendanceRecord.StatusGZ != true || attendanceRecord.StatusDO != true)
    //                    {
    //                        if (attendanceRecord.LateOut != null)
    //                        {
    //                            attendanceRecord.OTMin = attendanceRecord.LateOut;
    //                            attendanceRecord.StatusOT = true;
    //                            attendanceRecord.Remarks = attendanceRecord.Remarks + "[N-OT]";
    //                        }
    //                        else
    //                        {
    //                            attendanceRecord.Remarks.Replace("[OT]", "");
    //                            attendanceRecord.Remarks.Replace("[N-OT]", "");
    //                            attendanceRecord.OTMin = null;
    //                            attendanceRecord.StatusOT = null;
    //                        }
    //                    }
    //                }
    //            }




    //        }
    //        catch (Exception ex)
    //        {

    //        }
    //    }

    //    private void CalculateOpenShiftTimes(AttData attendanceRecord, Shift shift)
    //    {
    //        try
    //        {
    //            //Calculate WorkMin
    //            if (attendanceRecord != null)
    //            {
    //                if (attendanceRecord.TimeOut != null && attendanceRecord.TimeIn != null)
    //                {
    //                    attendanceRecord.Remarks = "";
    //                    TimeSpan mins = (TimeSpan)(attendanceRecord.TimeOut - attendanceRecord.TimeIn);
    //                    //Check if GZ holiday then place all WorkMin in GZOTMin
    //                    if (attendanceRecord.StatusGZ == true)
    //                    {
    //                        attendanceRecord.GZOTMin = (short)mins.TotalMinutes;
    //                        attendanceRecord.WorkMin = (short)mins.TotalMinutes;
    //                        attendanceRecord.StatusGZOT = true;
    //                        attendanceRecord.Remarks = attendanceRecord.Remarks + "[GZ-OT]";
    //                    }
    //                    else if (attendanceRecord.StatusDO == true)
    //                    {
    //                        attendanceRecord.OTMin = (short)mins.TotalMinutes;
    //                        attendanceRecord.WorkMin = (short)mins.TotalMinutes;
    //                        attendanceRecord.StatusOT = true;
    //                        attendanceRecord.Remarks = attendanceRecord.Remarks + "[R-OT]";
    //                    }
    //                    else
    //                    {
    //                        if (shift.HasBreak == true)
    //                        {
    //                            attendanceRecord.WorkMin = (short)(mins.TotalMinutes - shift.BreakMin);
    //                            attendanceRecord.ShifMin = (short)(CalculateShiftMinutes(shift, attendanceRecord.AttDate.Value.DayOfWeek) - (short)shift.BreakMin);
    //                        }
    //                        else
    //                        {
    //                            // CalculateShiftEndTime = ShiftStart + DutyHours
    //                            TimeSpan shiftEnd = CalculateShiftEndTime(shift, attendanceRecord.AttDate.Value.DayOfWeek);
    //                            attendanceRecord.WorkMin = (short)(mins.TotalMinutes);
    //                            //Calculate OverTIme, 
    //                            if (mins.TotalMinutes > (CalculateShiftMinutes(shift, attendanceRecord.AttDate.Value.DayOfWeek) + shift.OverTimeMin))
    //                            {
    //                                attendanceRecord.OTMin = (Int16)(Convert.ToInt16(mins.TotalMinutes) - CalculateShiftMinutes(shift, attendanceRecord.AttDate.Value.DayOfWeek));
    //                                attendanceRecord.WorkMin = (short)((mins.TotalMinutes) - attendanceRecord.OTMin);
    //                                attendanceRecord.StatusOT = true;
    //                                attendanceRecord.Remarks = attendanceRecord.Remarks + "[N-OT]";
    //                            }
    //                            //Calculate Early Out
    //                            if (mins.TotalMinutes < (CalculateShiftMinutes(shift, attendanceRecord.AttDate.Value.DayOfWeek) + shift.EarlyOut))
    //                            {
    //                                Int16 EarlyoutMin = (Int16)(Convert.ToInt16(mins.TotalMinutes) - CalculateShiftMinutes(shift, attendanceRecord.AttDate.Value.DayOfWeek));
    //                                if (EarlyoutMin > shift.EarlyOut)
    //                                {
    //                                    attendanceRecord.EarlyOut = EarlyoutMin;
    //                                    attendanceRecord.StatusEO = true;
    //                                    attendanceRecord.Remarks = attendanceRecord.Remarks + "[EO]";
    //                                }
    //                            }
    //                            // round off work mins if overtime less than shift.OverTimeMin >
    //                            if (attendanceRecord.WorkMin > CalculateShiftMinutes(shift, attendanceRecord.AttDate.Value.DayOfWeek) && (attendanceRecord.WorkMin <= (CalculateShiftMinutes(shift, attendanceRecord.AttDate.Value.DayOfWeek) + shift.OverTimeMin)))
    //                            {
    //                                attendanceRecord.WorkMin = CalculateShiftMinutes(shift, attendanceRecord.AttDate.Value.DayOfWeek);
    //                            }
    //                        }
    //                    }
    //                }
    //            }
    //        }
    //        catch (Exception ex)
    //        {
    //            string _error = "";
    //            if (ex.InnerException.Message != null)
    //                _error = ex.InnerException.Message;
    //            else
    //                _error = ex.Message;
    //            _myHelperClass.WriteToLogFile("Attendance Processing at Calculating Times;  " + _error);
    //        }
    //      }
    //}
}
