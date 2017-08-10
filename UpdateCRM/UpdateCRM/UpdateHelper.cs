using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using LinqToExcel;
using LinqToExcel.Attributes;
using System.Collections;

namespace UpdateHelper
{
    #region UpdateAcademics
    public class UpdateAcademics
    {
        //DECLARE XL COLUMNS TO READ
        [ExcelColumn("ContactID")]
        public string ContactID { get; set; }

        [ExcelColumn("Email")]
        public string Email { get; set; }

        [ExcelColumn ("Institution ID")]
        public string InstitutionID { get; set; }

        [ExcelColumn ("After Name")]
        public string AfterName { get; set; }

        [ExcelColumn("Selected Publications")]
        public string SelectedPublications { get; set; }

        [ExcelColumn("Interest")]
        public string Interest { get; set; }

        [ExcelColumn("ProExperience")]
        public string ProExperience { get; set; }

        [ExcelColumn ("Expertise")]
        public string Expertise { get; set; }

       

        //USE LINQTOEXCEL METHODS TO PARSE THE XL FILE
        public ExcelQueryFactory ReadAcademicData()
        {
         
            var academicExcel = new ExcelQueryFactory(@"C:\AutoNumber\CRM\191.xlsx")
            {
                DatabaseEngine = LinqToExcel.Domain.DatabaseEngine.Ace,
                TrimSpaces = LinqToExcel.Query.TrimSpacesType.Both,
                UsePersistentConnection = true,
                ReadOnly = true
            };

            return academicExcel;

            }
    
        }

  }

     
      


  

#endregion
