using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using TASDownloadService.Model;
using System.Net.NetworkInformation;

namespace TASDownloadService.Helper
{
    class Downloader
    {
        public void DownloadDataInIt()
        {
             List<ReadersCommLibrary.Poll> records = new List<ReadersCommLibrary.Poll>();
            using (var context = new TAS2013Entities())
            {
                List<Reader> readers = context.Readers.Where(aa=>aa.Status==true).ToList();
                foreach (var reader in readers)
                {
                    // Ping to Device and recieve true if device exists
                    if (IsConnectedToInternet(reader.IpAdd))
                    {
                        ReadersCommLibrary.IReader readerHelper = new ReadersCommLibrary.ZKReader();
                        //readerHelper.OnConnected += new EventHandler(OnConnectedDevice);

                        // Connect function and recieve true if connection established
                        if (readerHelper.Connect(reader.IpAdd, reader.IpPort))
                        {
                            try
                            {
                                // Download Records from Device and store it to records
                                SaveServiceLog(reader.RdrID, "Data Downloading Start", reader.RdrTypeID);
                                records = readerHelper.DownloadData(reader.RdrTypeID);
                                readerHelper.Disconnect();
                                // Save Device Attendance Data to Poll Data
                                try
                                {
                                    SaveAttDataToPollData(records, reader.RdrID);
                                    // Enter Service log -- ErrorCode = 5 -- Data Downloaded Complete
                                    SaveServiceLog(reader.RdrID, "Data Downloaded Complete-Total Records are:" + records.Count.ToString(), 5);
                                    // Enter Reader Event log of performed operation
                                    SaveRdrEventLog(reader.RdrID, "Download", 5);
                                }
                                catch (Exception ex)
                                {
                                    SaveServiceLog(reader.RdrID, "PollData Save Error with exception:" + ex.InnerException.Message, 6); //ErrorCode = 6 -- PollData Save Error
                                }  
                            }
                            catch (Exception e)
                            {
                                // Enter Service log if downloading not performed or recieve error
                                SaveServiceLog(reader.RdrID, "Data Downloaded Failed with exception:" + e.InnerException.Message, 4); //ErrorCode = 4 -- Data Download Failed

                            }

                        }
                        else
                        {
                            // enter Service log for Reader failed to Connect -- //ErrorCode = 2 -- notconnected
                            SaveServiceLog(reader.RdrID, "Reader failed to Connect", 2); 
                        }
                    }
                    else
                    {
                        //Service Log for reader not pinged
                        //ErrorCode = 1 -- notpinged
                        SaveServiceLog(reader.RdrID,"Reader Not Pinged",1);
                    }

                }
            }
            if (records.Count() > 0)
            {
                //Processor pr = new Processor();
                //pr.ProcessAttendance();
            }
        }

        private void OnConnectedDevice(object sender, EventArgs e)
        {
            
        }

        //public void DownloadData(List<Reader> reader)
        //{
        //    try
        //    {
        //        List<ReadersCommLibrary.Poll> records = new List<ReadersCommLibrary.Poll>();
        //        using (var context = new TAS2013Entities())
        //        {
        //            // Ping to Device and recieve true if device exists
        //            if (IsConnectedToInternet(reader.FirstOrDefault().IpAdd))
        //            {
        //                ReadersCommLibrary.IReader readerHelper = new ReadersCommLibrary.ZKReader();
        //                // Connect function and recieve true if connection established
        //                if (readerHelper.Connect(reader.FirstOrDefault().IpAdd, reader.FirstOrDefault().IpPort))
        //                {
        //                    try
        //                    {
        //                        // Download Records from Device and store it to records
        //                        SaveServiceLog(reader.FirstOrDefault().RdrID, "Data Downloading Start", 3);
        //                        records = readerHelper.DownloadData(5);
        //                        readerHelper.Disconnect();
        //                        // Save Device Attendance Data to Poll Data
        //                        try
        //                        {
        //                            SaveAttDataToPollData(records, reader.FirstOrDefault().RdrID);
        //                            // Enter Service log -- ErrorCode = 5 -- Data Downloaded Complete
        //                            SaveServiceLog(reader.FirstOrDefault().RdrID, "Data Downloaded Complete-Total Records are:" + records.Count.ToString(), 5);
        //                            //// Enter Reader Event log of performed operation
        //                            //SaveRdrEventLog(reader.FirstOrDefault().RdrID, "Download", 5);
        //                        }
        //                        catch (Exception ex)
        //                        {
        //                            SaveServiceLog(reader.FirstOrDefault().RdrID, "PollData Save Error with exception:" + ex.InnerException.Message, 6); //ErrorCode = 6 -- PollData Save Error
        //                        }
        //                    }
        //                    catch (Exception e)
        //                    {
        //                        // Enter Service log if downloading not performed or recieve error
        //                        SaveServiceLog(reader.FirstOrDefault().RdrID, "Data Downloaded Failed with exception:" + e.InnerException.Message, 4); //ErrorCode = 4 -- Data Download Failed

        //                    }

        //                }
        //                else
        //                {
        //                    // enter Service log for Reader failed to Connect -- //ErrorCode = 2 -- notconnected
        //                    SaveServiceLog(reader.FirstOrDefault().RdrID, "Reader failed to Connect", 2);
        //                }
        //            }
        //            else
        //            {
        //                //Service Log for reader not pinged
        //                //ErrorCode = 1 -- notpinged
        //                SaveServiceLog(reader.FirstOrDefault().RdrID, "Reader Not Pinged", 1);
        //            }
        //        }
        //    }
        //    catch (Exception ex)
        //    {

        //    }
            
        //}
        // Save Service Log
        private void SaveServiceLog(short _RdrID,string _Description,byte _ErrorCode)
        {
            try
            {
                using (var context = new TAS2013Entities())
                {
                    ServiceLog _ServiceLog = new ServiceLog();
                    _ServiceLog.Description = _Description;
                    _ServiceLog.ErrorCode = _ErrorCode;
                    _ServiceLog.DateTime = DateTime.Now;
                    _ServiceLog.ReaderID = _RdrID;
                    _ServiceLog.Processed = false;
                    context.ServiceLogs.AddObject(_ServiceLog);
                    context.SaveChanges();
                }
            }
            catch (Exception ex)
            {

            }
        }
        // Save Rdr Event log
        private void SaveRdrEventLog(short _RdrID, string _Description, byte _ErrorCode)
        {
            using (var context = new TAS2013Entities())
            {
                RdrEventLog _rdrEventLog = new RdrEventLog();
                _rdrEventLog.ReaderID = _RdrID;
                _rdrEventLog.Description = _Description;
                _rdrEventLog.ErrorCode = _ErrorCode;
                _rdrEventLog.EventDate = DateTime.Now;
                context.RdrEventLogs.AddObject(_rdrEventLog);
                context.SaveChanges();
            }
        }

        private void SaveAttDataToPollData(List<ReadersCommLibrary.Poll> records, Int16 _RdrID)
        {
            using (var context = new TAS2013Entities())
            {
                List<Emp> emps = context.Emps.Where(aa => aa.Status == true).ToList();
                foreach (var entry in records)
                {
                    try
                    {
                        var emp = emps.Where(aa => aa.FpID == entry.ID);
                        var reader = context.Readers.Where(aa => aa.RdrID == _RdrID);
                        if (emp.Count() > 0)
                        {
                            PollData p = new PollData();
                            p.EmpID = emp.FirstOrDefault().EmpID;
                            p.EntDate = entry.EntryDateTime.Date;
                            p.EntTime = entry.EntryDateTime;
                            p.EmpDate = emp.FirstOrDefault().EmpID.ToString() + entry.EntryDateTime.Date.ToString("yyMMdd");
                            p.CardNo = emp.FirstOrDefault().CardNo;
                            p.RdrID = reader.FirstOrDefault().RdrID;
                            p.FpID = entry.ID;
                            p.RdrDuty = reader.FirstOrDefault().RdrDutyCode.RdrDutyID;
                            p.Split = false;
                            p.Process = false;
                            p.AddedDate = DateTime.Today;
                            context.PollDatas.AddObject(p);
                            context.SaveChanges();
                        }
                        else
                        {
                            PollDataError pde = new PollDataError();
                            pde.DeviceRegID = entry.ID;
                            pde.EntryDate = entry.EntryDateTime.Date;
                            pde.EntryTime = entry.EntryDateTime;
                            pde.AddedDate = DateTime.Today;
                            context.PollDataErrors.AddObject(pde);
                            context.SaveChanges();
                        }
                    }
                    catch (Exception ex)
                    {

                    }
                }
            }
        }

        private bool IsConnectedToInternet(string _IPAddress)
        {
            //Uri url = new Uri(_IPAddress);
            //string pingurl = string.Format("{0}", url.Host);
            string host = _IPAddress;
            bool result = false;
            Ping p = new Ping();
            try
            {
                PingReply reply = p.Send(host, 3000);
                if (reply.Status == IPStatus.Success)
                    return true;
            }
            catch { }
            return result;
        }

        enum ErrorCode
        {
            NotPinged = 1,
            NotConnected = 2,
            DownloadStart = 3,
            DownloadFailed = 4,
            DownloadCompleted = 5,
            PollDataSaveError = 6
        }
    }
}
