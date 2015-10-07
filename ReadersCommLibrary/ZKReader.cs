using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace ReadersCommLibrary
{
    public class ZKReader : IReader
    {
        private zkemkeeper.CZKEMClass czkem;

        public ZKReader()
        {
            czkem = new zkemkeeper.CZKEMClass();
        }
        public bool Connect(string IPAddress, int Port)
        {
            bool result = false;
            try
            {
                if (czkem.Connect_Net(IPAddress, Port))
                    IsConnected = true;
                this.Port = Port;
                this.IPAddress = IPAddress;
                result = true;
            }
            catch (Exception)
            {
                throw;
            }
            return result;
        }

        public bool Disconnect()
        {
            bool result = false;
            try
            {
                czkem.Disconnect();
                IsConnected = false;
            }
            catch (Exception)
            {
                throw;
            }
            return result;
        }

        int iMachineNumber = 1;

        public List<Poll> DownloadData(int rdrType)
        {
            //int iMachineNumber = 1;
            string sdwEnrollNumber = "";
            int idwTMachineNumber = 0;
            int idwEMachineNumber = 0;
            int idwVerifyMode = 0;
            int idwInOutMode = 0;
            int idwYear = 0;
            int idwMonth = 0;
            int idwDay = 0;
            int idwHour = 0;
            int idwMinute = 0;
            int idwSecond = 0;
            int idwWorkcode = 0;
            int idwEnrollNumber = 0;
            int idwErrorCode = 0;
            int iGLCount = 0;
            int iIndex = 0;
            List<Poll> data = new List<Poll>();
            if (IsConnected)
            {
                czkem.EnableDevice(iMachineNumber, false);//disable the device
                if (czkem.ReadGeneralLogData(iMachineNumber))//read all the attendance records to the memory
                {
                    switch (rdrType)
                    {
                        case 7://Black and White Readers
                            while (czkem.GetGeneralLogData(iMachineNumber, ref idwTMachineNumber, ref idwEnrollNumber, ref idwEMachineNumber, ref idwVerifyMode, ref idwInOutMode, ref idwYear, ref idwMonth, ref idwDay, ref idwHour, ref idwMinute))//get records from the memory
                            {
                                try
                                {
                                    iGLCount++;
                                    try
                                    {
                                        DateTime entry = new DateTime(idwYear, idwMonth, idwDay, idwHour, idwMinute, idwSecond);
                                        data.Add(new Poll(Convert.ToInt32(idwEnrollNumber), entry));
                                        iIndex++;
                                    }
                                    catch (Exception ex)
                                    {

                                    }
                                }
                                catch (Exception ex)
                                {

                                }
                            }
                            break;
                        case 8://IFace
                            while (czkem.SSR_GetGeneralLogData(iMachineNumber, out sdwEnrollNumber, out idwVerifyMode, out idwInOutMode, out idwYear, out idwMonth, out idwDay, out idwHour, out idwMinute, out idwSecond, ref idwWorkcode))//get records from the memory
                            {
                                try
                                {
                                    iGLCount++;
                                    try
                                    {
                                        DateTime entry = new DateTime(idwYear, idwMonth, idwDay, idwHour, idwMinute, idwSecond);
                                        data.Add(new Poll(Convert.ToInt32(sdwEnrollNumber), entry));
                                        iIndex++;
                                    }
                                    catch (Exception ex)
                                    {

                                    }
                                }
                                catch (Exception ex)
                                {

                                }
                            }
                            break;
                    }
                }
                else
                {
                    czkem.GetLastError(ref idwErrorCode);
                }
                czkem.EnableDevice(iMachineNumber, true);//enable the device

            }
            if (data.Count > 0)
            {
                ClearAttDataFromDevice(iMachineNumber);
            }
            return data;
        }
        // Clear Attendance Data from Device
        public void ClearAttDataFromDevice(int _iMachineNum)
        {
            try
            {
                int idwErrorCode = 0;
                czkem.EnableDevice(_iMachineNum, false);//disable the device
                if (czkem.ClearGLog(_iMachineNum))
                {
                    czkem.RefreshData(_iMachineNum);//the data in the device should be refreshed
                }
                else
                {
                    czkem.GetLastError(ref idwErrorCode);
                }
                czkem.EnableDevice(iMachineNumber, true);//enable the device
            }
            catch (Exception ex)
            {
            }
        }

        public bool SetDateTime()
        {
            if (czkem.SetDeviceTime(iMachineNumber))
                return true;
            else
                return false;
        }

        string ip;
        int readerid;
        short dutycode;
        bool isConnected;
        int port;

        public string IPAddress
        {
            get
            {
                return ip;
            }
            set
            {
                ip = value;
            }
        }

        public int ReaderID
        {
            get
            {
                return readerid;
            }
            set
            {
                readerid = value;
            }
        }

        public int Port
        {
            get
            {
                return port;
            }
            set
            {
                port = value;
            }
        }

        public short DutyCode
        {
            get
            {
                return dutycode;
            }
            set
            {

                dutycode = value;
            }
        }

        public bool IsConnected
        {
            get
            {
                return isConnected;
            }
            set
            {
                isConnected = value;
            }
        }

        event EventHandler onconnectHandler;
        public event EventHandler OnConnected
        {
            add { this.onconnectHandler += value; }
            remove { this.onconnectHandler -= value; }
        }

        public event EventHandler OnDisconnected;

        void RaiseOnConnectedEvent(object sender, EventArgs args)
        {
            if (onconnectHandler != null)
            {
                onconnectHandler(sender, args);
            }
        }

    }

    }
