using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace themch.Models
{
    public class CauHoi
    {
        private static int maCH = 1;
        public CauHoi() { maCH++; }

       
        private string HinhAnh;
        private string NoiDubg;

        private List<DapAn> cauHois;
        public string HinhAnh1 { get => HinhAnh; set => HinhAnh = value; }
        public string NoiDubg1 { get => NoiDubg; set => NoiDubg = value; }
        public int MaCH { get => maCH; }
        public List<DapAn> CauHois1 { get => cauHois; set => cauHois = value; }
    }
}