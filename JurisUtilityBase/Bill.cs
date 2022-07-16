using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace JurisUtilityBase
{
    public class Bill
    {

        public Bill()
        {
            exps = new List<ExpAttachment>();

        }


        public int billNo { get; set; }

        public string clientNo { get; set; }

        public string clientName { get; set; }

        public string billDate { get; set; }

        public string matterNo { get; set; }

        public string matterName { get; set; }

        public int clisys { get; set; }

        public int matsys { get; set; }

        public bool badBill { get; set; }

        public bool hasExpAttach { get; set; }

        public List<ExpAttachment> exps { get; set; }



    }
}
