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
    class NoteAttachment
    {

        [ExcelColumn("RecordName")]
        public string RecordName { get; set; }

        [ExcelColumn("FileName")]
        public string FileName { get; set; }

        //Change these values as required

        //KYC - C:\FilesToUpload\KYC\KYC.xlsx Subject - KYC
        //Screnning Report - C:\FilesToUpload\ScreeningReport\ScreeningReport.xlsx Subject - Screening Report
        //MLRO - C:\FilesToUpload\MLRO\MLRO.xlsx Subject - MLRO
        //Certificate of Authority - C:\FilesToUpload\CertificateAuthority\CertificateAuthority.xlsx //Subject - CertificateOfIncorporation
        //Credit Reference - C:\FilesToUpload\CreditReference\CreditReference.xlsx //Subject - CertificateOfIncorporation
        //st120 - C:\FilesToUpload\ST120\st120.xlsx Subject - ST120

        string fileLocation = @"C:\FilesToUpload\ST120\st120.xlsx";
        string folderLocation = @"C:\FilesToUpload\ST120\";
        string subject = "ST120";
        
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

            var files = from p in fileValueExcel.Worksheet<NoteAttachment>("Values")
                        select p;

            foreach (var file in files)
            {

                string _recordName = file.RecordName;
                string _fileName = file.FileName;

                //Annotation(_fileName);
                GetRecordID(_recordName, _fileName);

            }
        }

        private void GetRecordID(string recordName, string fileName)
        {

            string connStr = ConfigurationManager.ConnectionStrings[1].ConnectionString;

            CrmServiceClient conn = new CrmServiceClient(connStr);

            // Cast the proxy client to the IOrganizationService interface.
            //_orgService = (IOrganizationService)conn.OrganizationWebProxyClient != null ? (IOrganizationService)conn.OrganizationWebProxyClient : (IOrganizationService)conn.OrganizationServiceProxy;
            IOrganizationService service = (IOrganizationService)conn.OrganizationServiceProxy;

            //Linq to retrieve Record Guid
            OrgServiceContext svc = new OrgServiceContext(service);

            var recordGuid = from r in svc.deb_applicationSet
                             where r.deb_name == recordName
                             select r;

            foreach (var r in recordGuid)
            {
                string _entity = r.LogicalName;
                string Id = r.Id.ToString();
                string _recordFirstName = recordName;
                Annotation(Id, fileName, _entity, recordName);
            }


        }

        private void Annotation(string recordId, string fileName, string entity, string recordName)
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

                Console.WriteLine("File attached to " + recordName);
            }
            
        }

    }

}


