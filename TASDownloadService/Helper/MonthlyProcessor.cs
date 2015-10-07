using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using TASDownloadService.Model;

namespace TASDownloadService.Helper
{
    //public class MonthlyProcessor
    //{
    //    TAS2013Entities context = new TAS2013Entities();
    //    List<AttData> EmpAttData = new List<AttData>();

    //    public bool processContractualMonthlyAttSingle(DateTime startDate, DateTime endDate, Emp _Emp, List<AttData> _EmpAttData)
    //    {
    //        string EmpMonth = _Emp.EmpID.ToString() + startDate.Date.Month.ToString();

    //        //Get Attendance data of employee according to selected month
    //        try
    //        {
    //            EmpAttData = _EmpAttData;
    //        }
    //        catch (Exception ex)
    //        {
    //            return false;
    //        }
    //        //Check for already processed data
    //        if (context.AttMnDatas.Where(aa => aa.EmpMonth == EmpMonth).Count() > 0)
    //        {
    //            AttMnData _TempAttMn = context.AttMnDatas.Where(aa => aa.EmpMonth == EmpMonth).FirstOrDefault();
    //            _attMonth = _TempAttMn;
    //        }
    //        TDays = Convert.ToByte((endDate - startDate).Days + 1);
    //        CalculateMonthlyAttendanceSheet(EmpAttData);
    //        _attMonth.Period = startDate.Date.Month.ToString() + startDate.Date.Year.ToString();
    //        _attMonth.EmpMonth = _Emp.EmpID.ToString() + startDate.Date.Month.ToString();
    //        _attMonth.EmpNo = _Emp.EmpNo;
    //        _attMonth.EmpID = _Emp.EmpID;
    //        _attMonth.EmpName = _Emp.EmpName;
    //        try
    //        {
    //            if (context.AttMnDatas.Where(aa => aa.EmpMonth == EmpMonth).Count() == 0)
    //            {
    //                context.AttMnDatas.AddObject(_attMonth);
    //            }
    //            context.SaveChanges();
    //        }
    //        catch (Exception ex)
    //        {
    //            return false;
    //        }
    //        return true;
    //    }

    //    AttMnData _attMonth = new AttMnData();
    //    byte TDays = 0;
    //    byte PresentDays = 0;
    //    byte AbsentDays = 0;
    //    byte LeaveDays = 0;
    //    byte RestDays = 0;
    //    byte GZDays = 0;
    //    Int16 EarlyIn = 0;
    //    Int16 EarlyOut = 0;
    //    Int16 LateIn = 0;
    //    Int16 LateOut = 0;
    //    Int16 WorkTime = 0;
    //    Int16 NOT = 0;
    //    Int16 GOT = 0;
    //    Int16 ExpectedWrkTime = 0;
    //    private void CalculateMonthlyAttendanceSheet(List<AttData> _EmpAttData)
    //    {
    //        foreach (var item in _EmpAttData)
    //        {
    //            //current day is GZ holiday
    //            if (item.StatusGZ == true || item.DutyCode == "G")
    //            {
    //                Marksheet(item.AttDate.Value.Day, "G");
    //                GZDays++;
    //            }
    //            //if current day is Rest day
    //            if (item.StatusDO == true || item.DutyCode == "R")
    //            {
    //                Marksheet(item.AttDate.Value.Day, "H");
    //                RestDays++;
    //            }
    //            //current day is leave
    //            if (item.DutyCode == "L" || item.StatusLeave == true)
    //            {
    //                Marksheet(item.AttDate.Value.Day, "L");
    //                LeaveDays++;
    //            }
    //            //if current day is absent
    //            if (item.StatusAB == true || item.DutyCode == "D")
    //            {
    //                Marksheet(item.AttDate.Value.Day, "A");
    //                AbsentDays++;
    //            }
    //            //currentday is present
    //            if (item.TimeIn != null && item.TimeOut != null)
    //            {
    //                if (item.DutyCode == "R" && item.DutyCode == "G")
    //                {

    //                }
    //                else
    //                {
    //                    Marksheet(item.AttDate.Value.Day, "P");
    //                    PresentDays++;
    //                }
    //            }
    //            try
    //            {
    //                //Sum EarlyIn/Out, LateIn/Out, WorkTime, NOT, GOT
    //                if (item.EarlyIn != null && item.EarlyIn > 0)
    //                    EarlyIn = (Int16)(EarlyIn + Convert.ToInt16(item.EarlyIn));
    //                if (item.EarlyOut != null && item.EarlyOut > 0)
    //                    EarlyOut = (Int16)(EarlyOut + Convert.ToInt16(item.EarlyOut));
    //                if (item.LateIn != null && item.LateIn > 0)
    //                    LateIn = (Int16)(LateIn + Convert.ToInt16(item.LateIn));
    //                if (item.LateOut != null && item.LateOut > 0)
    //                    LateOut = (Int16)(LateOut + Convert.ToInt16(item.LateOut));
    //                if (item.OTMin != null && item.OTMin > 0)
    //                    NOT = (Int16)(NOT + Convert.ToInt16(item.OTMin));
    //                if (item.GZOTMin != null && item.GZOTMin > 0)
    //                    GOT = (Int16)(GOT + Convert.ToInt16(item.GZOTMin));
    //                if (item.WorkMin != null && item.WorkMin > 0)
    //                    WorkTime = (Int16)(WorkTime + Convert.ToInt16(item.WorkMin));
    //                if (item.ShifMin != null && item.ShifMin > 0)
    //                    ExpectedWrkTime = (Int16)(ExpectedWrkTime + Convert.ToInt16(item.ShifMin));
    //            }
    //            catch (Exception ex)
    //            {

    //            }
    //        }
    //        //
    //        _attMonth.TotalDays = TDays;
    //        _attMonth.PreDays = PresentDays;
    //        _attMonth.AbDays = AbsentDays;
    //        _attMonth.LeaveDays = LeaveDays;
    //        _attMonth.RestDays = RestDays;
    //        _attMonth.GZDays = GZDays;
    //        _attMonth.WorkDays = (byte)(PresentDays + AbsentDays + LeaveDays);

    //        _attMonth.TEarlyIn = EarlyIn;
    //        _attMonth.TEarlyOut = EarlyOut;
    //        _attMonth.TLateIn = LateIn;
    //        _attMonth.TLateOut = LateOut;
    //        _attMonth.TWorkTime = WorkTime;
    //        _attMonth.TGZOT = GOT;
    //        _attMonth.TNOT = NOT;
    //        _attMonth.ExpectedWrkTime = ExpectedWrkTime;
    //    }

    //    public void Marksheet(int day, string _Code)
    //    {
    //        switch (day)
    //        {
    //            case 1:
    //                _attMonth.D1 = _Code;
    //                break;
    //            case 2:
    //                _attMonth.D2 = _Code;
    //                break;
    //            case 3:
    //                _attMonth.D3 = _Code;
    //                break;
    //            case 4:
    //                _attMonth.D4 = _Code;
    //                break;
    //            case 5:
    //                _attMonth.D5 = _Code;
    //                break;
    //            case 6:
    //                _attMonth.D6 = _Code;
    //                break;
    //            case 7:
    //                _attMonth.D7 = _Code;
    //                break;
    //            case 8:
    //                _attMonth.D8 = _Code;
    //                break;
    //            case 9:
    //                _attMonth.D9 = _Code;
    //                break;
    //            case 10:
    //                _attMonth.D10 = _Code;
    //                break;
    //            case 11:
    //                _attMonth.D11 = _Code;
    //                break;
    //            case 12:
    //                _attMonth.D12 = _Code;
    //                break;
    //            case 13:
    //                _attMonth.D13 = _Code;
    //                break;
    //            case 14:
    //                _attMonth.D14 = _Code;
    //                break;
    //            case 15:
    //                _attMonth.D15 = _Code;
    //                break;
    //            case 16:
    //                _attMonth.D16 = _Code;
    //                break;
    //            case 17:
    //                _attMonth.D17 = _Code;
    //                break;
    //            case 18:
    //                _attMonth.D18 = _Code;
    //                break;
    //            case 19:
    //                _attMonth.D19 = _Code;
    //                break;
    //            case 20:
    //                _attMonth.D20 = _Code;
    //                break;
    //            case 21:
    //                _attMonth.D21 = _Code;
    //                break;
    //            case 22:
    //                _attMonth.D22 = _Code;
    //                break;
    //            case 23:
    //                _attMonth.D23 = _Code;
    //                break;
    //            case 24:
    //                _attMonth.D24 = _Code;
    //                break;
    //            case 25:
    //                _attMonth.D25 = _Code;
    //                break;
    //            case 26:
    //                _attMonth.D26 = _Code;
    //                break;
    //            case 27:
    //                _attMonth.D27 = _Code;
    //                break;
    //            case 28:
    //                _attMonth.D28 = _Code;
    //                break;
    //            case 29:
    //                _attMonth.D29 = _Code;
    //                break;
    //            case 30:
    //                _attMonth.D30 = _Code;
    //                break;
    //            case 31:
    //                _attMonth.D31 = _Code;
    //                break;
    //        }
    //    }
    //}
}
