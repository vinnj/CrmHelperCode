using System;
using System.Linq;
using System.Collections.Generic;
using LinqToExcel;
using LinqToExcel.Attributes;
using UpdateCRM;

namespace UpdateCRM
{

    public class Update 
    {
        static void Main(string[] args)
        {
            Academics academics = new Academics();
            academics.RetrieveUpdatedAcademics();

        }
    }
}
