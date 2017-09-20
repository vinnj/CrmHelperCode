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

namespace NoteAttachment
{
    internal class AgreementNoteAttachment
    {
        //Change these values as required

        //KYC - C:\FilesToUpload\KYC\KYC.xlsx
        //Screnning Report - C:\FilesToUpload\ScreeningReport\ScreeningReport.xlsx
        //MLRO - C:\FilesToUpload\MLRO\MLRO.xlsx
        //Agreement - C:\FilesToUpload\Agreement.Agreement.xlsx
        //Certificate of Authority - C:\FilesToUpload\CertificateAuthority\CertificateAuthority.xlsx
        //Credit Reference - C:\FilesToUpload\CreditReference\CreditReference.xlsx
        //Contact ID proof - C:\FilesToUpload\Id\ContactId.xlsx

        //EntityReference entityRef = new EntityReference();

        private readonly string fileLocation = @"C:\FilesToUpload\Agreement\Agreement.xlsx";
        private readonly string folderLocation = @"C:\FilesToUpload\Agreement\";
        private readonly string subject = "Agreement";

        [ExcelColumn("RecordName")]
        public string RecordName { get; set; }

        [ExcelColumn("FileName")]
        public string FileName { get; set; }

        [ExcelColumn("Account")]
        public string Account { get; set; }

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

            var files = from p in fileValueExcel.Worksheet<AgreementNoteAttachment>("Values")
                select p;

            foreach (var file in files)
            {
                var _recordName = file.RecordName;
                var _fileName = file.FileName;
                var _accountId = new EntityReference("account", new Guid(file.Account));

                //Annotation(_fileName);
                GetRecordID(_recordName, _fileName, _accountId);
            }
        }

        private void GetRecordID(string recordName, string fileName, EntityReference accountID)
        {
            var connStr = ConfigurationManager.ConnectionStrings[1].ConnectionString;

            var conn = new CrmServiceClient(connStr);

            // Cast the proxy client to the IOrganizationService interface.
            //_orgService = (IOrganizationService)conn.OrganizationWebProxyClient != null ? (IOrganizationService)conn.OrganizationWebProxyClient : (IOrganizationService)conn.OrganizationServiceProxy;
            IOrganizationService service = conn.OrganizationServiceProxy;

            //Linq to retrieve Record Guid
            var svc = new OrgServiceContext(service);

            var recordGuid = from r in svc.deb_agreementSet
                where r.deb_name == recordName &&
                      r.deb_AccountId == accountID
                select r;

            foreach (var r in recordGuid)
            {
                var _entity = r.LogicalName;
                var Id = r.Id.ToString();
                var _accountName = r.deb_AccountId.Name;

                Annotation(Id, fileName, _entity, recordName, _accountName);
            }
        }

        private void Annotation(string recordId, string fileName, string entity, string recordName, string accountName)
        {
            if (fileName != null)
            {
                var connStr = ConfigurationManager.ConnectionStrings[1].ConnectionString;

                var conn = new CrmServiceClient(connStr);

                IOrganizationService service = conn.OrganizationServiceProxy;


                var _stream = File.OpenRead(folderLocation + fileName);
                var _bData = new byte[_stream.Length];
                _stream.Read(_bData, 0, _bData.Length);
                _stream.Close();
                var encodedData = Convert.ToBase64String(_bData);


                var _annotation = new Entity("annotation");
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