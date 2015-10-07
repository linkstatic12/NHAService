using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using TASDownloadService.Model;

namespace TASDownloadService.AttProcessSummary
{
    public class DailySummaryClass
    {
        public static void PreviousTenDaysSummary(DateTime dateStart, String criteria, int criteriaValue)
        {
            //for (int i = 0; i < 10; i++)
            //{
            //    DateTime date = dateStart.AddDays(-i);
            //    CalculateSummary(date, criteria, criteriaValue);
            //}
        }

        public DailySummaryClass(DateTime dateStart, DateTime dateEnd)
        {
            List<Region> regions = new List<Region>();
            List<Location> locs = new List<Location>();
            List<Section> secs = new List<Section>();
            List<Department> wings = new List<Department>();
            regions = db.Regions.ToList();
            while (dateStart <= dateEnd)
            {
                CalculateSummary(dateStart, "C", 1);
                foreach (var region in regions)
                {
                    CalculateSummary(dateStart, "R", region.RegionID);
                }
                dateStart = dateStart.AddDays(1);
            }
        }
        TAS2013Entities db = new TAS2013Entities();
        public static void CalculateSummary(DateTime dateStart, String criteria, int criteriaValue)
        {
            bool ProcssedDS = false;
            DateTime makeMe = new DateTime(dateStart.Year, dateStart.Month, dateStart.Day, 0, 0, 0);
            SummaryEntity summary = new SummaryEntity();

            TAS2013Entities context = new TAS2013Entities();
            ViewMultipleInOut vmio = new ViewMultipleInOut();
            summary.SummaryDateCriteria = dateStart.ToString("yyMMdd") + summary.criterianame + criteriaValue;
            String day = DateTime.Today.DayOfWeek.ToString();
            day = day.Substring(0, 3);
            day = day + "Min";

            List<ViewMultipleInOut> attList = new List<ViewMultipleInOut>();
            switch (criteria)
            {
                case "C": attList = context.ViewMultipleInOuts.Where(aa => aa.AttDate == dateStart).ToList();
                    //Change the below line for NHA please
                    summary.criterianame = context.Options.FirstOrDefault().CompanyName;
                    summary.SummaryDateCriteria = summary.SummaryDateCriteria + criteria + criteriaValue;

                    break;

                case "R": attList = context.ViewMultipleInOuts.Where(aa => aa.AttDate == dateStart).ToList();
                    attList = attList.Where(aa => aa.RegionID == criteriaValue).ToList();
                    if (attList.Count > 0)
                        summary.criterianame = attList[0].RegionName;
                    break;
                case "D": attList = context.ViewMultipleInOuts.Where(aa => aa.AttDate == dateStart && aa.DeptID == criteriaValue).ToList();
                    if (attList.Count > 0)
                        summary.criterianame = attList[0].DeptName;
                    break;
                case "E": attList = context.ViewMultipleInOuts.Where(aa => aa.AttDate == dateStart && aa.SecID == criteriaValue).ToList();
                    if (attList.Count > 0)
                        summary.criterianame = attList[0].SectionName;
                    break;
                case "L": attList = context.ViewMultipleInOuts.Where(aa => aa.AttDate == dateStart && aa.LocID == criteriaValue).ToList();
                    if (attList.Count > 0)
                        summary.criterianame = attList[0].LocName;
                    break;
                case "S": attList = context.ViewMultipleInOuts.Where(aa => aa.AttDate == dateStart && aa.ShiftID == criteriaValue).ToList();
                    if (attList.Count > 0)
                        summary.criterianame = attList[0].ShiftName;
                    break;
                default: attList = context.ViewMultipleInOuts.Where(aa => aa.AttDate >= dateStart).ToList();
                    if (attList.Count > 0)
                        summary.criterianame = context.Options.FirstOrDefault().CompanyName;
                    break;



            }
            if (attList.Count > 0)
            {
                DailySummary dailysumm = new DailySummary();
                if (context.DailySummaries.Where(aa => aa.SummaryDateCriteria == summary.SummaryDateCriteria).Count() > 0)
                {
                    dailysumm = context.DailySummaries.First(aa => aa.SummaryDateCriteria == summary.SummaryDateCriteria);
                    ProcssedDS = true;
                }
                dailysumm.SummaryDateCriteria = summary.SummaryDateCriteria;
                dailysumm.Criteria = criteria;
                dailysumm.CriteriaValue = (short)criteriaValue;
                dailysumm.CriteriaName = summary.criterianame;
                dailysumm.TotalEmps = (short)attList.Count;
                dailysumm.PresentEmps = (short)attList.Where(aa => aa.StatusP == true).Count<ViewMultipleInOut>();
                dailysumm.AbsentEmps = (short)(dailysumm.TotalEmps - dailysumm.PresentEmps);
                dailysumm.Date = makeMe;
                dailysumm.EIEmps = (short)attList.Where(aa => aa.StatusEI == true).Count<ViewMultipleInOut>();
                dailysumm.EOEmps = (short)attList.Where(aa => aa.StatusEO == true).Count<ViewMultipleInOut>();
                dailysumm.LIEmps = (short)attList.Where(aa => aa.StatusLI == true).Count<ViewMultipleInOut>();
                dailysumm.LOEmps = (short)attList.Where(aa => aa.StatusLO == true).Count<ViewMultipleInOut>();
                dailysumm.LvEmps = 0;
                dailysumm.HalfLvEmps = 0;
                dailysumm.ShortLvEmps = 0;
                dailysumm.ExpectedWorkMins = 0;
                dailysumm.ActualWorkMins = 0;
                dailysumm.LossWorkMins = 0;
                dailysumm.OTMins = 0;
                dailysumm.LIMins = 0;
                dailysumm.LOMins = 0;
                dailysumm.EIMins = 0;
                dailysumm.EOMins = 0;
                dailysumm.AOTMins = 0;
                dailysumm.OTEmps = 0;
                dailysumm.OTMins = 0;
                long averageTimeIn = 0;
                long averageTimeOut = 0;
                //LV ,short half
                foreach (var emp in attList)
                {
                    if (emp.TimeIn != null)
                        averageTimeIn = averageTimeIn + emp.TimeIn.Value.Ticks;
                    if (emp.TimeOut != null)
                        averageTimeOut = averageTimeOut + emp.TimeOut.Value.Ticks;

                    if (emp.LateIn != null)
                        dailysumm.LIMins = dailysumm.LIMins + emp.LateIn;
                    if (emp.LateOut != null)
                        dailysumm.LOMins = dailysumm.LOMins + emp.LateOut;
                    if (emp.EarlyIn != null)
                        dailysumm.EIMins = dailysumm.EIMins + emp.EarlyIn;
                    if (emp.EarlyOut != null)
                        dailysumm.EOMins = dailysumm.EOMins + emp.EarlyOut;
                    if (emp.WorkMin != null)
                        dailysumm.ActualWorkMins = dailysumm.ActualWorkMins + emp.WorkMin;
                    //code leave bundle
                    if (emp.StatusLeave == true)
                        dailysumm.LvEmps = dailysumm.LvEmps++;
                    if (emp.StatusHL == true)
                        dailysumm.HalfLvEmps = dailysumm.HalfLvEmps++;
                    if (emp.StatusSL == true)
                        dailysumm.ShortLvEmps = dailysumm.ShortLvEmps++;

                    //code for over time emps
                    if (emp.StatusOT == true)
                    {
                        dailysumm.OTEmps = (short)(dailysumm.OTEmps + 1);
                        dailysumm.OTMins = dailysumm.OTMins + emp.OTMin;
                    }

                    //code for day off
                    if (emp.StatusDO == true)
                        dailysumm.DayOffEmps++;


                    //code for expected work mins
                    if (emp.ShifMin != null)
                        dailysumm.ExpectedWorkMins = emp.ShifMin + dailysumm.ExpectedWorkMins;




                }


                // dailysumm.OnTimeEmps = (short)(dailysumm.TotalEmps - dailysumm.OTEmps);

                dailysumm.DayOffEmps = (short)(dailysumm.ShortLvEmps + dailysumm.LvEmps);
                dailysumm.LossWorkMins = dailysumm.ExpectedWorkMins - dailysumm.ActualWorkMins;

                try
                {
                    if (dailysumm.PresentEmps != 0)
                        dailysumm.AvgTimeIn = new DateTime((long)(averageTimeIn / dailysumm.PresentEmps)).TimeOfDay;
                    if (dailysumm.PresentEmps != 0)
                        dailysumm.AvgTimeOut = new DateTime((long)(averageTimeOut / dailysumm.PresentEmps)).TimeOfDay;
                    if (dailysumm.TotalEmps != 0)
                        dailysumm.AActualMins = (short)(dailysumm.ActualWorkMins / dailysumm.TotalEmps);
                    if (dailysumm.TotalEmps != 0)
                        dailysumm.ALossMins = (short)(dailysumm.LossWorkMins / dailysumm.TotalEmps);
                    if (dailysumm.TotalEmps != 0)
                        dailysumm.AExpectedMins = (short)(dailysumm.ExpectedWorkMins / dailysumm.TotalEmps);
                    if (dailysumm.LIEmps != 0)
                        dailysumm.ALIMins = (short)(dailysumm.LIMins / dailysumm.LIEmps);
                    if (dailysumm.LOEmps != 0)
                        dailysumm.ALOMins = (short)(dailysumm.LOMins / dailysumm.LOEmps);
                    if (dailysumm.EIEmps != 0)
                        dailysumm.AEIMins = (short)(dailysumm.EIMins / dailysumm.EIEmps);
                    if (dailysumm.EOEmps != 0)
                        dailysumm.AEOMins = (short)(dailysumm.EOMins / dailysumm.EOEmps);
                    if (dailysumm.OTEmps != 0)
                        dailysumm.AOTMins = (short)(dailysumm.OTMins / dailysumm.OTEmps);
                }
                catch (Exception e)
                { }
                if (ProcssedDS == false)
                    context.DailySummaries.AddObject(dailysumm);
                context.SaveChanges();



            }


        }

    }
}
