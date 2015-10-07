using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReadersCommLibrary
{
    public interface IReader
    {
        bool Connect(string IPAddress, int Port);
        bool Disconnect();
        List<Poll> DownloadData(int rdrType);
        bool SetDateTime();
        event EventHandler OnConnected;
        event EventHandler OnDisconnected;
        string IPAddress { get; set; }
        int ReaderID { get; set; }
        short DutyCode { get; set; }
        int Port { get; set; }
        bool IsConnected { get; set; }

    }
}
