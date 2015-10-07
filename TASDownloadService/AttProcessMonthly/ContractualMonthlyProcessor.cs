using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using WMSFFService;
using TASDownloadService.Model;

namespace WMSFFService
{
    public class ContractualMonthlyProcessor
    {
        TAS2013Entities context = new TAS2013Entities();
        List<AttData> EmpAttData = new List<AttData>();
        MyCustomFunctions _myHelperClass = new MyCustomFunctions();
        public bool processContractualMonthlyAttSingle(DateTime startDate, DateTime endDate, Emp _Emp, List<AttData> _EmpAttData)
        {
            string EmpMonth = _Emp.EmpID + startDate.Date.Month.ToString();
            //Attendance data of employee for selected dates
            try
            {
                EmpAttData = _EmpAttData;
            }
            catch (Exception ex)
            {
                string _error = "";
                if (ex.InnerException.Message != null)
                    _error = ex.InnerException.Message;
                else
                    _error = ex.Message;
                _myHelperClass.WriteToLogFile("Exception at Contactual Monthly Process: " + _error);

                return false;
            }
            //Check for already processed data
            if (context.AttMnDatas.Where(aa => aa.EmpMonth == EmpMonth).Count() > 0)
            {
                AttMnData _TempAttMn = context.AttMnDatas.Where(aa => aa.EmpMonth == EmpMonth).FirstOrDefault();
                _attMonth = _TempAttMn;
                _attMonth.OT1 = 0;
                _attMonth.OT2 = 0;
                _attMonth.OT3 = 0;
                _attMonth.OT4 = 0;
                _attMonth.OT5 = 0;
                _attMonth.OT6 = 0;
                _attMonth.OT7 = 0;
                _attMonth.OT8 = 0;
                _attMonth.OT9 = 0;
                _attMonth.OT10 = 0;
                _attMonth.OT11 = 0;
                _attMonth.OT12 = 0;
                _attMonth.OT13 = 0;
                _attMonth.OT14 = 0;
                _attMonth.OT15 = 0;
                _attMonth.OT16 = 0;
                _attMonth.OT17 = 0;
                _attMonth.OT18 = 0;
                _attMonth.OT19 = 0;
                _attMonth.OT20 = 0;
                _attMonth.OT21 = 0;
                _attMonth.OT22 = 0;
                _attMonth.OT23 = 0;
                _attMonth.OT24 = 0;
                _attMonth.OT25 = 0;
                _attMonth.OT26 = 0;
                _attMonth.OT27 = 0;
                _attMonth.OT28 = 0;
                _attMonth.OT29 = 0;
                _attMonth.OT30 = 0;
                _attMonth.OT31 = 0;
                _attMonth.L1 = null;
                _attMonth.L2 = null;
                _attMonth.L3 = null;
                _attMonth.L4 = null;
                _attMonth.L5 = null;
                _attMonth.L6 = null;
                _attMonth.L7 = null;
                _attMonth.L8 = null;
                _attMonth.L9 = null;
                _attMonth.L10 = null;
                _attMonth.L11 = null;
                _attMonth.L12 = null;
                _attMonth.L13 = null;
                _attMonth.L14 = null;
                _attMonth.L15 = null;
                _attMonth.L16 = null;
                _attMonth.L17 = null;
                _attMonth.L18 = null;
                _attMonth.L19 = null;
                _attMonth.L20 = null;
                _attMonth.L21 = null;
                _attMonth.L22 = null;
                _attMonth.L23 = null;
                _attMonth.L24 = null;
                _attMonth.L25 = null;
                _attMonth.L26 = null;
                _attMonth.L27 = null;
                _attMonth.L28 = null;
                _attMonth.L29 = null;
                _attMonth.L30 = null;
                _attMonth.L31 = null;

            }

            TDays = Convert.ToByte((endDate - startDate).Days + 1);
            List<LvData> _lvData = context.LvDatas.Where(aa => aa.AttDate >= startDate && aa.AttDate <= endDate).ToList();
            CalculateMonthlyAttendanceSheet(EmpAttData, _lvData);
            _attMonth.Period = startDate.Date.Month.ToString() + startDate.Date.Year.ToString();
            _attMonth.EmpMonth = _Emp.EmpID + startDate.Date.Month.ToString();
            _attMonth.EmpNo = _Emp.EmpNo;
            _attMonth.EmpID = _Emp.EmpID;
            _attMonth.EmpName = _Emp.EmpName;
            try
            {
                if (context.AttMnDatas.Where(aa => aa.EmpMonth == EmpMonth).Count() == 0)
                {
                    context.AttMnDatas.AddObject(_attMonth);
                }
                context.SaveChanges();
            }
            catch (Exception ex)
            {
                string _error = "";
                if (ex.InnerException.Message != null)
                    _error = ex.InnerException.Message;
                else
                    _error = ex.Message;
                _myHelperClass.WriteToLogFile("Exception at Contactual Monthly Process 2: " + _error);
                return false;
            }
            return true;
        }

        AttMnData _attMonth = new AttMnData();
        byte TDays = 0;
        float PresentDays = 0;
        float AbsentDays = 0;
        byte LeaveDays = 0;
        float HalfleaveDays = 0;
        byte RestDays = 0;
        byte GZDays = 0;
        Int16 EarlyIn = 0;
        Int16 EarlyOut = 0;
        Int16 LateIn = 0;
        Int16 WorkTime = 0;
        Int16 NOT = 0;
        Int16 GOT = 0;
        private void CalculateMonthlyAttendanceSheet(List<AttData> _EmpAttData, List<LvData> _LvData)
        {

            foreach (var item in _EmpAttData)
            {
                Int16 OverTime = 0;
                //current day is GZ holiday
                if (item.StatusGZ == true && item.DutyCode == "G")
                {
                    Marksheet(item.AttDate.Value.Day, "G");
                    GZDays++;
                    if (item.GZOTMin != null && item.GZOTMin > 5)
                    {
                        MarkOverTime(item.AttDate.Value.Day, Convert.ToInt16(item.GZOTMin));
                    }
                }
                //if current day is Rest day
                if (item.StatusDO == true && item.DutyCode == "R")
                {
                    Marksheet(item.AttDate.Value.Day, "R");
                    RestDays++;
                    if (item.OTMin != null && item.OTMin > 5)
                    {
                        MarkOverTime(item.AttDate.Value.Day, Convert.ToInt16(item.OTMin));
                    }
                }


                //current day is leave
                if (item.StatusLeave == true)
                {
                    if (item.TimeIn == null && item.TimeOut == null)
                    {
                        Marksheet(item.AttDate.Value.Day, "L");
                        string LvType = _LvData.Where(aa => aa.EmpDate == item.EmpDate && aa.HalfLeave != true).FirstOrDefault().LvCode;
                        string LvTypeName = "";
                        switch (LvType)
                        {
                            case "A"://Caasual
                                LvTypeName = "CL";
                                break;
                            case "B"://Anual
                                LvTypeName = "AL";
                                break;
                            case "C":
                                LvTypeName = "SL";
                                break;
                        }
                        MarkLeaveType(item.AttDate.Value.Day, LvTypeName);
                        LeaveDays++;
                    }
                }
                //if current day is absent
                if (item.StatusAB == true && item.DutyCode == "D")
                {
                    Marksheet(item.AttDate.Value.Day, "A");
                    AbsentDays++;
                }
                //currentday is present
                if (item.TimeIn != null && item.TimeOut != null)
                {
                    if (item.DutyCode=="D")
                    {
                        Marksheet(item.AttDate.Value.Day, "P");
                        if (item.StatusDO != true && item.StatusGZ != true)
                        {
                            PresentDays++;
                        }
                        if (item.OTMin != null && item.OTMin > 0)
                        {
                            OverTime = (Int16)(OverTime + Convert.ToInt16(item.OTMin));
                        }
                        if (item.OTMin > 0)
                        {
                            MarkOverTime(item.AttDate.Value.Day, Convert.ToInt16(item.OTMin));
                        }
                        if (item.GZOTMin > 0)
                        {
                            MarkOverTime(item.AttDate.Value.Day, Convert.ToInt16(item.GZOTMin));
                        }
                    }
                }
                ////current day is Half Leave
                if (item.StatusHL == true)
                {
                    if (item.TimeIn != null && item.TimeOut != null)
                    {
                        Marksheet(item.AttDate.Value.Day, "H");
                        string LvType = _LvData.Where(aa => aa.EmpDate == item.EmpDate && aa.HalfLeave == true).FirstOrDefault().LvCode;
                        string LvTypeName = "";
                        switch (LvType)
                        {
                            case "A"://Caasual
                                LvTypeName = "HCL";
                                break;
                            case "B"://Anual
                                LvTypeName = "HAL";
                                break;
                            case "C":
                                LvTypeName = "HSL";
                                break;
                        }
                        MarkLeaveType(item.AttDate.Value.Day, LvTypeName);
                        HalfleaveDays = (float)(HalfleaveDays + 1);
                    }
                    else
                    {
                        HalfleaveDays = (float)(HalfleaveDays + 1);
                        AbsentDays = (float)(AbsentDays - 0.5);
                    }
                }
                //Manual 
                if (item.StatusMN == true && item.DutyCode == "D")
                {
                    if (item.TimeIn == null && item.TimeOut == null)
                    {
                        if (item.StatusP == true)
                        {
                            if (item.StatusDO != true && item.StatusAB != true)
                            {
                                Marksheet(item.AttDate.Value.Day, "P");
                                PresentDays++;
                            }
                        }
                    }
                }
                //Missing Attendance
                if ((item.TimeIn == null && item.TimeOut != null) || (item.TimeIn != null && item.TimeOut == null))
                {
                    Marksheet(item.AttDate.Value.Day, "I");
                }
                try
                {
                    //Sum EarlyIn/Out, LateIn/Out, WorkTime, NOT, GOT
                    if (item.EarlyIn != null && item.EarlyIn > 5)
                        EarlyIn = (Int16)(EarlyIn + Convert.ToInt16(item.EarlyIn));
                    if (item.EarlyOut != null && item.EarlyOut > 5)
                        EarlyOut = (Int16)(EarlyOut + Convert.ToInt16(item.EarlyOut));
                    if (item.LateIn != null && item.LateIn > 5)
                        LateIn = (Int16)(LateIn + Convert.ToInt16(item.LateIn));
                    if (item.OTMin != null && item.OTMin > 5)
                        NOT = (Int16)(NOT + Convert.ToInt16(item.OTMin));
                    if (item.GZOTMin != null && item.GZOTMin > 5)
                        GOT = (Int16)(GOT + Convert.ToInt16(item.GZOTMin));
                    if (item.WorkMin != null && item.WorkMin > 5)
                        WorkTime = (Int16)(WorkTime + Convert.ToInt16(item.WorkMin));
                }
                catch (Exception ex)
                {

                }
            }
            //
            PresentDays = PresentDays - (float)(HalfleaveDays / 2);
            _attMonth.TotalDays = TDays;
            _attMonth.PreDays = PresentDays;
            _attMonth.AbDays = AbsentDays;
            _attMonth.LeaveDays = LeaveDays;
            _attMonth.RestDays = RestDays;
            _attMonth.GZDays = GZDays;
            _attMonth.WorkDays = (float)(PresentDays + RestDays + GZDays + LeaveDays + (HalfleaveDays / 2));
            _attMonth.HalfLeavesDay = HalfleaveDays;
            _attMonth.TEarlyIn = EarlyIn;
            _attMonth.TEarlyOut = EarlyOut;
            _attMonth.TLateIn = LateIn;
            _attMonth.TWorkTime = WorkTime;
            _attMonth.TGZOT = GOT;
            _attMonth.TNOT = NOT;
        }

        private void MarkLeaveType(int day, string LeaveType)
        {
            switch (day)
            {
                case 1:
                    _attMonth.L1 = LeaveType;
                    break;
                case 2:
                    _attMonth.L2 = LeaveType;
                    break;
                case 3:
                    _attMonth.L3 = LeaveType;
                    break;
                case 4:
                    _attMonth.L4 = LeaveType;
                    break;
                case 5:
                    _attMonth.L5 = LeaveType;
                    break;
                case 6:
                    _attMonth.L6 = LeaveType;
                    break;
                case 7:
                    _attMonth.L7 = LeaveType;
                    break;
                case 8:
                    _attMonth.L8 = LeaveType;
                    break;
                case 9:
                    _attMonth.L9 = LeaveType;
                    break;
                case 10:
                    _attMonth.L10 = LeaveType;
                    break;
                case 11:
                    _attMonth.L11 = LeaveType;
                    break;
                case 12:
                    _attMonth.L12 = LeaveType;
                    break;
                case 13:
                    _attMonth.L13 = LeaveType;
                    break;
                case 14:
                    _attMonth.L14 = LeaveType;
                    break;
                case 15:
                    _attMonth.L15 = LeaveType;
                    break;
                case 16:
                    _attMonth.L16 = LeaveType;
                    break;
                case 17:
                    _attMonth.L17 = LeaveType;
                    break;
                case 18:
                    _attMonth.L18 = LeaveType;
                    break;
                case 19:
                    _attMonth.L19 = LeaveType;
                    break;
                case 20:
                    _attMonth.L20 = LeaveType;
                    break;
                case 21:
                    _attMonth.L21 = LeaveType;
                    break;
                case 22:
                    _attMonth.L22 = LeaveType;
                    break;
                case 23:
                    _attMonth.L23 = LeaveType;
                    break;
                case 24:
                    _attMonth.L24 = LeaveType;
                    break;
                case 25:
                    _attMonth.L25 = LeaveType;
                    break;
                case 26:
                    _attMonth.L26 = LeaveType;
                    break;
                case 27:
                    _attMonth.L27 = LeaveType;
                    break;
                case 28:
                    _attMonth.L28 = LeaveType;
                    break;
                case 29:
                    _attMonth.L29 = LeaveType;
                    break;
                case 30:
                    _attMonth.L30 = LeaveType;
                    break;
                case 31:
                    _attMonth.L31 = LeaveType;
                    break;
            }
        }

        public void Marksheet(int day, string _Code)
        {
            switch (day)
            {
                case 1:
                    _attMonth.D1 = _Code;
                    break;
                case 2:
                    _attMonth.D2 = _Code;
                    break;
                case 3:
                    _attMonth.D3 = _Code;
                    break;
                case 4:
                    _attMonth.D4 = _Code;
                    break;
                case 5:
                    _attMonth.D5 = _Code;
                    break;
                case 6:
                    _attMonth.D6 = _Code;
                    break;
                case 7:
                    _attMonth.D7 = _Code;
                    break;
                case 8:
                    _attMonth.D8 = _Code;
                    break;
                case 9:
                    _attMonth.D9 = _Code;
                    break;
                case 10:
                    _attMonth.D10 = _Code;
                    break;
                case 11:
                    _attMonth.D11 = _Code;
                    break;
                case 12:
                    _attMonth.D12 = _Code;
                    break;
                case 13:
                    _attMonth.D13 = _Code;
                    break;
                case 14:
                    _attMonth.D14 = _Code;
                    break;
                case 15:
                    _attMonth.D15 = _Code;
                    break;
                case 16:
                    _attMonth.D16 = _Code;
                    break;
                case 17:
                    _attMonth.D17 = _Code;
                    break;
                case 18:
                    _attMonth.D18 = _Code;
                    break;
                case 19:
                    _attMonth.D19 = _Code;
                    break;
                case 20:
                    _attMonth.D20 = _Code;
                    break;
                case 21:
                    _attMonth.D21 = _Code;
                    break;
                case 22:
                    _attMonth.D22 = _Code;
                    break;
                case 23:
                    _attMonth.D23 = _Code;
                    break;
                case 24:
                    _attMonth.D24 = _Code;
                    break;
                case 25:
                    _attMonth.D25 = _Code;
                    break;
                case 26:
                    _attMonth.D26 = _Code;
                    break;
                case 27:
                    _attMonth.D27 = _Code;
                    break;
                case 28:
                    _attMonth.D28 = _Code;
                    break;
                case 29:
                    _attMonth.D29 = _Code;
                    break;
                case 30:
                    _attMonth.D30 = _Code;
                    break;
                case 31:
                    _attMonth.D31 = _Code;
                    break;
            }
        }

        //For OT
        public void MarkOverTime(int day, Int16 _OTMin)
        {
            switch (day)
            {
                case 1:
                    _attMonth.OT1 = _OTMin;
                    break;
                case 2:
                    _attMonth.OT2 = _OTMin;
                    break;
                case 3:
                    _attMonth.OT3 = _OTMin;
                    break;
                case 4:
                    _attMonth.OT4 = _OTMin;
                    break;
                case 5:
                    _attMonth.OT5 = _OTMin;
                    break;
                case 6:
                    _attMonth.OT6 = _OTMin;
                    break;
                case 7:
                    _attMonth.OT7 = _OTMin;
                    break;
                case 8:
                    _attMonth.OT8 = _OTMin;
                    break;
                case 9:
                    _attMonth.OT9 = _OTMin;
                    break;
                case 10:
                    _attMonth.OT10 = _OTMin;
                    break;
                case 11:
                    _attMonth.OT11 = _OTMin;
                    break;
                case 12:
                    _attMonth.OT12 = _OTMin;
                    break;
                case 13:
                    _attMonth.OT13 = _OTMin;
                    break;
                case 14:
                    _attMonth.OT14 = _OTMin;
                    break;
                case 15:
                    _attMonth.OT15 = _OTMin;
                    break;
                case 16:
                    _attMonth.OT16 = _OTMin;
                    break;
                case 17:
                    _attMonth.OT17 = _OTMin;
                    break;
                case 18:
                    _attMonth.OT18 = _OTMin;
                    break;
                case 19:
                    _attMonth.OT19 = _OTMin;
                    break;
                case 20:
                    _attMonth.OT20 = _OTMin;
                    break;
                case 21:
                    _attMonth.OT21 = _OTMin;
                    break;
                case 22:
                    _attMonth.OT22 = _OTMin;
                    break;
                case 23:
                    _attMonth.OT23 = _OTMin;
                    break;
                case 24:
                    _attMonth.OT24 = _OTMin;
                    break;
                case 25:
                    _attMonth.OT25 = _OTMin;
                    break;
                case 26:
                    _attMonth.OT26 = _OTMin;
                    break;
                case 27:
                    _attMonth.OT27 = _OTMin;
                    break;
                case 28:
                    _attMonth.OT28 = _OTMin;
                    break;
                case 29:
                    _attMonth.OT29 = _OTMin;
                    break;
                case 30:
                    _attMonth.OT30 = _OTMin;
                    break;
                case 31:
                    _attMonth.OT31 = _OTMin;
                    break;
            }
        }
    }
}
