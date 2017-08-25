using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Tooling;
using Microsoft.Xrm.Tooling.Connector;
using System.Configuration;
using System.IO;
using LinqToExcel;
using LinqToExcel.Attributes;
using Debeers.Xrm.Entities;

namespace NoteAttachment
{
    class AgreementNoteAttachment
    {
        [ExcelColumn("RecordName")]
        public string RecordName { get; set; }

        [ExcelColumn("FileName")]
        public string FileName { get; set; }
        
        [ExcelColumn("Account")]
        public string Account { get; set; }

        //Change these values as required

        //KYC - C:\FilesToUpload\KYC\KYC.xlsx
        //Screnning Report - C:\FilesToUpload\ScreeningReport\ScreeningReport.xlsx
        //MLRO - C:\FilesToUpload\MLRO\MLRO.xlsx
        //Agreement - C:\FilesToUpload\Agreement.Agreement.xlsx
        //Certificate of Authority - C:\FilesToUpload\CertificateAuthority\CertificateAuthority.xlsx
        //Credit Reference - C:\FilesToUpload\CreditReference\CreditReference.xlsx
        //Contact ID proof - C:\FilesToUpload\Id\ContactId.xlsx

        //EntityReference entityRef = new EntityReference();

        string fileLocation = @"C:\FilesToUpload\Agreement\Agreement.xlsx";
        string folderLocation = @"C:\FilesToUpload\Agreement\";
        string subject = "Agreement";

        private ExcelQueryFactory ReadData()
        {

            var fileUploadValues = new ExcelQueryFactory(fileLocation)
            {
                DatabaseEngine = LinqToExcel.Domain.DatabaseEngine.Ace,
                TrimSpaces = LinqToExcel.Query.TrimSpacesType.Both,
                UsePersistentConnection = true,
                ReadOnly = true
            };

            return fileUploadValues;

        }

        public void PassValues()
        {
            var fileValueExcel = ReadData();

            var files = from p in fileValueExcel.Worksheet<AgreementNoteAttachment>("Values")
                        select p;

            foreach (var file in files)
            {

                string _recordName = file.RecordName;
                string _fileName = file.FileName;
                EntityReference _accountId = new EntityReference("account", new Guid(file.Account));

                //Annotation(_fileName);
                GetRecordID(_recordName, _fileName, _accountId);

            }
        }

        private void GetRecordID(string recordName, string fileName, EntityReference accountID)
        {

            string connStr = ConfigurationManager.ConnectionStrings[1].ConnectionString;

            CrmServiceClient conn = new CrmServiceClient(connStr);

            // Cast the proxy client to the IOrganizationService interface.
            //_orgService = (IOrganizationService)conn.OrganizationWebProxyClient != null ? (IOrganizationService)conn.OrganizationWebProxyClient : (IOrganizationService)conn.OrganizationServiceProxy;
            IOrganizationService service = (IOrganizationService)conn.OrganizationServiceProxy;

            //Linq to retrieve Record Guid
            OrgServiceContext svc = new OrgServiceContext(service);

            var recordGuid = from r in svc.deb_agreementSet
                             where r.deb_name == recordName &&
                             r.deb_AccountId == accountID
                             select r;

            foreach (var r in recordGuid)
            {
                string _entity = r.LogicalName;
                string Id = r.Id.ToString();
                string _accountName = r.deb_AccountId.Name;

                Annotation(Id, fileName, _entity, recordName, _accountName);
            }


        }

        private void Annotation(string recordId, string fileName, string entity, string recordName, string accountName)
        {
            if(fileName != null)
            {
                string connStr = ConfigurationManager.ConnectionStrings[1].ConnectionString;

                CrmServiceClient conn = new CrmServiceClient(connStr);

                IOrganizationService service = (IOrganizationService)conn.OrganizationServiceProxy;


                FileStream _stream = File.OpenRead(folderLocation + fileName);
                byte[] _bData = new byte[_stream.Length];
                _stream.Read(_bData, 0, _bData.Length);
                _stream.Close();
                string encodedData = System.Convert.ToBase64String(_bData);


                Entity _annotation = new Entity("annotation");
                _annotation.Attributes["objectid"] = new EntityReference(entity, new Guid(recordId));
                _annotation.Attributes["objecttypecode"] = entity;
                _annotation.Attributes["subject"] = subject;
                _annotation.Attributes["documentbody"] = encodedData;
                _annotation.Attributes["mimetype"] = @"application/pdf";
                //_annotation.Attributes["notetext"] = "Credit Reference";
                _annotation.Attributes["filename"] = fileName;

                service.Create(_annotation);

                Console.WriteLine("File attached to " + accountName);
            }
           
        }

    }
}

