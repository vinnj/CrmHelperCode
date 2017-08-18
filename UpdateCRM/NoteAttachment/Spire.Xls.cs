using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Spire.Xls;
using System.Data;
using System.IO;
using LinqToExcel;

namespace NoteAttachment
{
    class Spire
    {
        public void MergeFiles()
        {
            Workbook workbook = new Workbook();
            //load the first workbook

            workbook.LoadFromFile(@"merge1.xlsx");
            //load the second workbook

            Workbook workbook2 = new Workbook();
            workbook2.LoadFromFile(@"merge2.xlsx");

            //import the second workbook's worksheet into the first workbook using a datatable
            Worksheet sheet2 = workbook2.Worksheets[0];
            DataTable dataTable = sheet2.ExportDataTable();

            Worksheet sheet1 = workbook.Worksheets[0];
            sheet1.InsertDataTable(dataTable, false, sheet1.LastRow + 1, 1);


            //save the workbook
            workbook.SaveToFile("result.xlsx");
        }

    }
}


