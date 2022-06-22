﻿/*
 * Company instance containing info from the Company tab in Quickbook
 */

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace QBToT4PDF
{
    public class Company
    {
        public string name { get; set; }
        public string addressBlock { get; set; }
        public string addressFull { get; set; }
        public string phone { get; set; }

        public Company()
        {
            name = "";
            addressBlock = "";
            addressFull = "";
            phone = "";
        }
         

    }
}
