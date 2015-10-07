using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ReadersCommLibrary
{
    public class EmpFPTemp
    {
        public Int32 FPID { get; set; }
        public string FP1 { get; set; }
        public string FP2 { get; set; }
        public string FP3 { get; set; }
        public string FP4 { get; set; }


        public EmpFPTemp(Int32 FPID, string _FP1, string _FP2, string _FP3, string _FP4)
        {
            this.FPID = FPID;
             this.FP1 = string.Copy(_FP1);
            this.FP2 = _FP2;
            this.FP3 = _FP3;
            this.FP4 = _FP4;
        }
    }
}
