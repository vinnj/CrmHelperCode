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
using NoteAttachment;
using Spire.Xls;

namespace NoteAttachment
{
   
    class ContactNoteAttachment
    {

        [ExcelColumn("RecordFirstName")]
        public string RecordFirstName { get; set; }

        [ExcelColumn("FileName")]
        public string FileName { get; set; }

        [ExcelColumn("RecordLastName")]
        public string RecordLastName { get; set; }

        //Contact ID proof - C:\FilesToUpload\Id\ContactId.xlsx

        //Change these values as required
        string fileLocation = @"C:\FilesToUpload\Id\ContactId.xlsx";
        string folderLocation = @"C:\FilesToUpload\Id\";
        string subject = "Identification";

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

        private void PassValues()
        {
            var fileValueExcel = ReadData();

            var files = from p in fileValueExcel.Worksheet<ContactNoteAttachment>("Values")
                        select p;

            foreach (var file in files)
            {
                
                string _recordFirstName = file.RecordFirstName;
                string _recordLastName = file.RecordLastName;
                string _fileName = file.FileName;

                //Annotation(_fileName);
                GetRecordID(_recordFirstName,_recordLastName,_fileName);
              
            }
        }

        private void GetRecordID(string recordFirstName,string recordLastName, string fileName)
        {
           
            string connStr = ConfigurationManager.ConnectionStrings[1].ConnectionString;

            CrmServiceClient conn = new CrmServiceClient(connStr);

            // Cast the proxy client to the IOrganizationService interface.
            //_orgService = (IOrganizationService)conn.OrganizationWebProxyClient != null ? (IOrganizationService)conn.OrganizationWebProxyClient : (IOrganizationService)conn.OrganizationServiceProxy;
            IOrganizationService service = (IOrganizationService)conn.OrganizationServiceProxy;

            //Linq to retrieve Record Guid
            OrgServiceContext svc = new OrgServiceContext(service);

            var recordGuid = from r in svc.ContactSet
                             where r.FirstName == recordFirstName &&
                             r.LastName == recordLastName
                             select r;

            foreach (var r in recordGuid)
            {
                string _entity = r.LogicalName;
                string Id = r.Id.ToString();
                string _recordFirstName = recordFirstName;
                Annotation(Id, fileName, _entity, recordFirstName);
            }
        }
        private void Annotation(string recordId, string fileName, string entity, string recordFirstName)
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

            Console.WriteLine("File attached to " + recordFirstName);
            
        }
        static void Main(string[] args)
        {
            //ContactNoteAttachment c = new ContactNoteAttachment();
            //c.PassValues();

            //NoteAttachment n = new NoteAttachment();
            //n.PassValues();

            //AgreementNoteAttachment a = new AgreementNoteAttachment();
            //a.PassValues();

            AccountNoteAttachment acc = new AccountNoteAttachment();
            acc.PassValues();


        
        }
    }
}
