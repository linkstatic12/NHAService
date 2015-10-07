using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using System.Threading;
using TASDownloadService.Model;
using System.Net.NetworkInformation;
using TASDownloadService.Helper;
using WMSFFService;
using TASDownloadService.AttProcessSummary;
using TASDownloadService.AttProcessDaily;

namespace TASDownloadService
{
    public partial class TASService : ServiceBase
    {
        System.Timers.Timer timer;
        static bool isTimerRunning = false;
        Thread DownloadThread = null;
        private static bool isServiceRunning = false;
        TimeSpan ProcessTime = new TimeSpan(16, 20, 00);
        public TASService()
        {
            InitializeComponent();
            timer = new System.Timers.Timer();
            timer.Interval = 5000;
            timer.AutoReset = true;
            timer.Elapsed += new System.Timers.ElapsedEventHandler(timer_Elapsed);
        }
        MyCustomFunctions _myHelperClass = new MyCustomFunctions();
        #region -- Service Start, Stop --
        // Service Start
        protected override void OnStart(string[] args)
        {
            try
            {
                timer.Start();
                DownloadThread = new Thread(RunService);
                Thread.Sleep(2000);
                DownloadThread.Start();
                isServiceRunning = true;
                _myHelperClass.WriteToLogFile("******************WMS Service Started*************** " );
            }
            catch (Exception ex)
            {

            }
        }

        protected override void OnStop()
        {
            try
            {
                isServiceRunning = false;
                timer.Stop();
                DownloadThread.Join();
                _myHelperClass.WriteToLogFile("******************WMS Service Stopped*************** ");
            }
            catch (Exception ex)
            {

            }
        }
        public void StartService()
        {
            OnStart(null);
        }
        public void StopService()
        {
            OnStop();
        }
        #endregion


        private void RunService()
        {
            while (isServiceRunning)
            {

            }
        }
        void timer_Elapsed(object sender, System.Timers.ElapsedEventArgs e)
        {
            try
            {
                timer.Stop();
                using (var context = new TAS2013Entities())
                {
                    // If Service download data on specific times
                    if (context.Options.FirstOrDefault().DownTime == true)
                    {
                        List<DownloadTime> _downloadTime = new List<DownloadTime>();
                        _downloadTime = context.DownloadTimes.ToList();
                        foreach (var item in _downloadTime)
                        {
                            // Add 2 minutes extra to download time
                            TimeSpan _dwnTimeEnd = (TimeSpan)item.DownTime + new TimeSpan(0, 2, 0);
                            if (DateTime.Now.TimeOfDay >= item.DownTime && DateTime.Now.TimeOfDay <= _dwnTimeEnd)
                            {
                                GlobalSettings._dateTime = Properties.Settings.Default.MyDateTime;
                                //Download Data from Devices
                                DownloadDataFromDevices();
                                //set Process = 1 where Date varies from ranges
                                AdjustPollData();
                                //Prcoess PollData to Attendance Data
                                ProcessAttendance pa = new ProcessAttendance();
                                pa.ProcessDailyAttendance();
                                DailySummaryClass dailysum = new DailySummaryClass(DateTime.Today.AddDays(-20), DateTime.Today);
                            }
                            /////////Process Monthly Attendance//////////
                            TimeSpan monthlyTStart = new TimeSpan(18, 50, 0);
                            TimeSpan monthlyTEnd = new TimeSpan(18, 52, 0);
                            if (DateTime.Now.TimeOfDay >= monthlyTStart && DateTime.Now.TimeOfDay <= monthlyTEnd)
                            {
                                DailySummaryClass dailysum = new DailySummaryClass(DateTime.Today.AddDays(-20),DateTime.Today );
                                //DailySummaryClass dailysum = new DailySummaryClass(new DateTime(2015, 07, 27), new DateTime(2015, 07, 27));
                                ////Correct Flags for monthly
                                DateTime dtStart = DateTime.Today.AddDays(-2);
                                DateTime dtend = DateTime.Today;
                                CorrectAttEntriesWithWrongFlags(dtStart, dtend);
                                ProcessMonthlyAttendance();
                            }
                        }
                    }
                    else
                    {
                    }

                }

            }
            catch
            {

            }
            finally
            {
                timer.Start();
            }
        }
        private void ProcessManualAttendance()
        {
            using (var ctx = new TAS2013Entities())
            {
                ManualProcess mp = new ManualProcess();
                List<Emp> emps = new List<Emp>();
                DateTime date = new DateTime(2015, 09, 10);
                emps = ctx.Emps.Where(aa => aa.Status == true).ToList();
                //emps.AddRange(ctx.Emps.Where(aa => (aa.EmpType.CatID == 2 && aa.CompanyID == 1)).ToList());
                //emps.AddRange(ctx.Emps.Where(aa => (aa.EmpType.CatID == 4) && aa.CompanyID == 1).ToList());
                //emps.AddRange(ctx.Emps.Where(aa => (aa.EmpType.CatID == 1) && aa.CompanyID == 1).ToList());
                //mp.BootstrapAttendance(date, emps);
                List<AttData> atts = new List<AttData>();
                //atts.AddRange(ctx.AttDatas.Where(aa => aa.Emp.EmpType.CatID == 2 && aa.Emp.CompanyID == 1 && aa.AttDate == date));
                //atts.AddRange(ctx.AttDatas.Where(aa => aa.Emp.EmpType.CatID == 4 && aa.Emp.CompanyID == 1 && aa.AttDate == date));
                //atts.AddRange(ctx.AttDatas.Where(aa => aa.Emp.EmpType.CatID == 1 && aa.Emp.CompanyID == 1 && aa.AttDate == date));
                atts = ctx.AttDatas.Where(aa => aa.AttDate == date).ToList();
                mp.ManualProcessAttendance(date, emps, atts);
            }
        }
        private void AdjustPollData()
        {
            int currentYear = DateTime.Today.Date.Year;
            int StartYear = currentYear - 3;
            DateTime endDate = DateTime.Today.AddDays(2);
            DateTime startDate = new DateTime(StartYear,1,1);
            using (var ctx = new TAS2013Entities())
            {
                List<PollData> polls = ctx.PollDatas.Where(aa=>aa.EntDate <= startDate &&aa.Process == false).ToList();
                foreach (var item in polls)
                {
                    item.Process = true;
                }
                ctx.SaveChanges();
                polls.Clear();
                polls = ctx.PollDatas.Where(aa => aa.EntDate >= endDate && aa.Process == false).ToList();
                foreach (var item in polls)
                {
                    item.Process = true;
                }
                ctx.SaveChanges();
                ctx.Dispose();
            }
        }

        private void CorrectAttEntriesWithWrongFlags(DateTime startdate, DateTime endDate)
        {
            using (var ctx = new TAS2013Entities())
            {
                // where StatusGZ ==1 and DutyCode != G
                List<AttData> _attDataForGZ = ctx.AttDatas.Where(aa => aa.AttDate >= startdate && aa.AttDate<= endDate && aa.StatusGZ==true && aa.DutyCode != "G").ToList();
                foreach (var item in _attDataForGZ)
                {
                    item.DutyCode  ="G";
                }
                ctx.SaveChanges();

                // where StatusDO ==1 and DutyCode != R
                List<AttData> _attDataForDO = ctx.AttDatas.Where(aa => aa.AttDate >= startdate && aa.AttDate <= endDate && aa.StatusDO == true && aa.DutyCode != "R").ToList();
                foreach (var item in _attDataForGZ)
                {
                    item.DutyCode = "R";
                }
                ctx.SaveChanges();
                ctx.Dispose();
            }
        }
        //Process Monthly Attendance for Contractuals
        private void ProcessMonthlyAttendance()
        {
            //Process Month till end of month
            DateTime endDate = DateTime.Today.Date;
            int currentDay = DateTime.Today.Date.Day;
            int currentMonth = DateTime.Today.Date.Month;
            int currentYear = DateTime.Today.Date.Year;
            DateTime startDate = new DateTime(currentYear, currentMonth, 1);
            if (currentMonth == 1&& currentDay <10)
            {
                currentMonth = 13;
                currentYear = currentYear - 1;
            }
            if (endDate.Day < 10)
            {
                currentMonth = currentMonth - 1;
                int DaysInPreviousMonth = System.DateTime.DaysInMonth(currentYear, currentMonth);
                endDate = new DateTime(currentYear, currentMonth, DaysInPreviousMonth);
                startDate = new DateTime(currentYear, currentMonth, 1);
            }
            else
            {
                startDate = new DateTime(currentYear, currentMonth, 1);
            }

            
            ProcessMonthly(startDate, endDate);
        }
        //Process Contractuals Monthly Attendance
        private void ProcessMonthly(DateTime startDate, DateTime endDate)
        {
            TAS2013Entities ctx = new TAS2013Entities();
            // Pass list of selected emp Attendance data to optimize sql query 
            List<AttData> _AttData = new List<AttData>();
            List<AttData> _EmpAttData = new List<AttData>();
            _AttData = ctx.AttDatas.Where(aa => aa.AttDate >= startDate && aa.AttDate <= endDate).ToList();
            int count = 0;
            List<Emp> _Emp = ctx.Emps.Where(em => em.Status == true).ToList();
            List<Emp> _oEmp = ctx.Emps.Where(em => em.Status == true).ToList();
            _Emp.AddRange(_oEmp);
            int _TE = _Emp.Count;
            foreach (Emp emp in _Emp)
            {
                count++;
                try
                {
                    ContractualMonthlyProcessor cmp = new ContractualMonthlyProcessor();
                    _EmpAttData = _AttData.Where(aa => aa.EmpID == emp.EmpID).ToList();
                    if (!cmp.processContractualMonthlyAttSingle(startDate, endDate, emp, _EmpAttData))
                    {
                    }

                }
                catch (Exception ex)
                {

                }
            }
        }

        private void DownloadDataFromDevices()
        {
            try
            {
                Downloader d = new Downloader();
                d.DownloadDataInIt();
            }
            catch (Exception ex)
            {

            }
        }

    }
}
