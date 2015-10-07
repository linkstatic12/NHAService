using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using WMSFFService;
using TASDownloadService.Model;

namespace WMSFFService
{
    class CreateMissingAttendance
    {
        public void CreatemissingAttendance(DateTime _startDate, DateTime _endDate)
        {

            using (var ctx = new TAS2013Entities())
            {
                List<AttData> _AttData = new List<AttData>();
                List<Emp> _Emp = new List<Emp>();
                DateTime CurrentDate = _startDate;
                _Emp = ctx.Emps.Where(aa => aa.Status == true).ToList();
                while (CurrentDate <= _endDate)
                {
                    _AttData = ctx.AttDatas.Where(aa => aa.AttDate == CurrentDate).ToList();
                    foreach (var emp in _Emp)
                    {
                        if (_AttData.Where(aa => aa.EmpID == emp.EmpID).Count() == 0)
                        {
                            //Create Attendance
                            CreateAttendance(CurrentDate, _Emp.Where(aa => aa.EmpID == emp.EmpID).ToList());
                        }
                    }
                    CurrentDate = CurrentDate.AddDays(1);
                }
            }
        }

        private void CreateAttendance(DateTime dateTime, List<Emp> _emp)
        {
            using (var ctx = new TAS2013Entities())
            {
                List<Roster> _Roster = new List<Roster>();
                _Roster = ctx.Rosters.Where(aa => aa.RosterDate == dateTime).ToList();
                List<RosterDetail> _NewRoster = new List<RosterDetail>();
                _NewRoster = ctx.RosterDetails.Where(aa => aa.RosterDate == dateTime).ToList();
                List<LvData> _LvData= new List<LvData>();
                _LvData = ctx.LvDatas.Where(aa => aa.AttDate == dateTime).ToList();
                List<LvShort> _lvShort = new List<LvShort>();
                _lvShort = ctx.LvShorts.Where(aa => aa.DutyDate == dateTime).ToList();
                List<AttData> _AttData = ctx.AttDatas.Where(aa => aa.AttDate == dateTime).ToList();
                List<Holiday> _Holidays = ctx.Holidays.ToList();
                foreach (var emp in _emp)
                {
                    string empDate = emp.EmpID + dateTime.ToString("yyMMdd");
                    if (_AttData.Where(aa => aa.EmpDate == empDate).Count() == 0)
                    {
                        try
                        {
                            /////////////////////////////////////////////////////
                            //  Mark Everyone Absent while creating Attendance //
                            /////////////////////////////////////////////////////
                            //Set DUTYCODE = D, StatusAB = true, and Remarks = [Absent]
                            AttData att = new AttData();
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
                            //  Check for GZ Day //
                            ////////////////////////
                            //Set DutyCode = R, StatusAB=false, StatusGZ = true, and Remarks=[GZ]
                            if (emp.Shift.GZDays == true)
                            {
                                foreach (var holiday in _Holidays)
                                {
                                    if (_Holidays.Where(hol => hol.HolDate.Month == att.AttDate.Value.Month && hol.HolDate.Day == att.AttDate.Value.Day).Count() > 0)
                                    {
                                        att.DutyCode = "G";
                                        att.StatusAB = false;
                                        att.StatusGZ = true;
                                        att.Remarks = "[GZ]";
                                        att.ShifMin = 0;
                                    }
                                }
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
                                }
                                else
                                {
                                    att.ShifMin = roster.WorkMin;
                                    att.DutyCode = "D";
                                    att.DutyTime = roster.DutyTime;
                                }
                            }
                            else if(_NewRoster.Where(aa => aa.CriteriaValueDate == crewCdate).Count() > 0)
                            {
                                var roster = _NewRoster.FirstOrDefault(aa => aa.CriteriaValueDate == crewCdate);
                                if (roster.WorkMin == 0)
                                {
                                    att.StatusAB = false;
                                    att.StatusDO = true;
                                    att.Remarks = "[DO]";
                                    att.DutyCode = "R";
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
                                }
                                else
                                {
                                    att.ShifMin = roster.WorkMin;
                                    att.DutyCode = "D";
                                    att.DutyTime = roster.DutyTime;
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
                                        att.Remarks = "[Casual Leave]";
                                    else if (Leave.LvCode == "B")
                                        att.Remarks = "[Annual Leave]";
                                    else if (Leave.LvCode == "C")
                                        att.Remarks = "[Sick Leave]";
                                    else
                                        att.Remarks = "[Leave]";
                                }
                                else
                                {
                                    att.StatusLeave = false;
                                }
                            }

                            /////////////////////////
                            //Check for Half Leave///
                            ////////////////////////
                            var _HalfLeave = _LvData.Where(lv => lv.EmpDate == att.EmpDate && lv.HalfLeave == true);
                            if (_HalfLeave.Count() > 0)
                            {
                                att.StatusLeave = true;
                                att.StatusAB = false;
                                att.DutyCode = "L";
                                att.StatusDO = false;
                                att.StatusHL = true;
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
                                att.StatusHL = false;
                            }
                            ctx.AttDatas.AddObject(att);
                            ctx.SaveChanges();
                        }
                        catch (Exception ex)
                        {
                            _myHelperClass.WriteToLogFile("Error In Creating Attendance of Employee: " + emp.EmpNo);
                        }
                    }
                }
            }
            _myHelperClass.WriteToLogFile("Creating Attendance of Date: "+dateTime.ToString());
        }
        MyCustomFunctions _myHelperClass = new MyCustomFunctions();
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
    }
}
