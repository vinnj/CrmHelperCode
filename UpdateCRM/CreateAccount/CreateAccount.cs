using System;
using System.Configuration;
using System.IO;
using System.Linq;
using Debeers.Xrm.Entities;
using LinqToExcel;
using LinqToExcel.Attributes;
using LinqToExcel.Domain;
using LinqToExcel.Query;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Tooling.Connector;


namespace CreateAccount
{
    public class CreateAccount
    {
        [ExcelColumn("Account Name")]
        public string accountName { get; set; }

        [ExcelColumn("Company Legal Type")]
        public string companyLegalType { get; set; }

        [ExcelColumn("Company ID Type")]
        public string companyIdType { get; set; }

        private readonly string fileLocation = @"C:\FileToLoad\DebeersAccount.xlsx";
        // private readonly string folderLocation = @"C:\FilesToUpload\Agreement\";
        // private readonly string subject = "Agreement";

        private ExcelQueryFactory ReadData()
        {
            var fileUploadValues = new ExcelQueryFactory(fileLocation)
            {
                DatabaseEngine = DatabaseEngine.Ace,
                TrimSpaces = TrimSpacesType.Both,
                UsePersistentConnection = true,
                ReadOnly = true
            };

            return fileUploadValues;
        }

        public void PassValues()
        {
            var fileValueExcel = ReadData();

            var files = from p in fileValueExcel.Worksheet<CreateAccount>("Account")
                select p;

            foreach (var file in files)
            {
                var _accountName = file.accountName;
                var _companyId = file.companyIdType;
                var _companyLegalType = file.companyLegalType;

                //Annotation(_fileName);
                CreateCrmAccount(_accountName, _companyId, _companyLegalType);
            }
        }


        public void CreateCrmAccount(string accountName, string companyIdType, string companyLegalType)
        {
            var connStr = ConfigurationManager.ConnectionStrings[1].ConnectionString;

            var conn = new CrmServiceClient(connStr);

            // Cast the proxy client to the IOrganizationService interface.
            //_orgService = (IOrganizationService)conn.OrganizationWebProxyClient != null ? (IOrganizationService)conn.OrganizationWebProxyClient : (IOrganizationService)conn.OrganizationServiceProxy;
            IOrganizationService service = conn.OrganizationServiceProxy;

            Guid companyKindId = Guid.Empty;
            int idType = 0;
           
            if (companyIdType == "EIN")
            {
                idType = 277470001;
                
            }

            if (companyIdType == "Tax ID")
            {
                idType = 277470000;
            }

            var svc = new OrgServiceContext(service);

            var query = from c in svc.deb_companykindSet
                where c.deb_name == companyLegalType
                select c;
            foreach (var c in query)
            {
                companyKindId = c.Id;
                
            }

            Account application = new Account
            {
               Name = accountName,
               deb_IDType = new OptionSetValue(idType),
               deb_companykind = new EntityReference("deb_companykind", companyKindId)
            };

            service.Create(application);
            Console.WriteLine("Account {0} created ", accountName);

        }

        static void Main(string[] args)
        {
            CreateAccount a = new CreateAccount();
            a.PassValues();
        }
    }
}
