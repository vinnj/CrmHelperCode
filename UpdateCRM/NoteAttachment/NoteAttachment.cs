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

namespace NoteAttachment
{
    class NoteAttachment
    {
        private void Annotation()
        {
           
            string connStr = ConfigurationManager.ConnectionStrings[1].ConnectionString;

            CrmServiceClient conn = new CrmServiceClient(connStr);

            // Cast the proxy client to the IOrganizationService interface.
            //_orgService = (IOrganizationService)conn.OrganizationWebProxyClient != null ? (IOrganizationService)conn.OrganizationWebProxyClient : (IOrganizationService)conn.OrganizationServiceProxy;
            IOrganizationService service = (IOrganizationService)conn.OrganizationServiceProxy;

            string filePath = @"C:\Users\vinujd\Documents\TestAttachment\sample.txt";
            byte[] fileContent = File.ReadAllBytes(filePath);
            string encodedData = System.Convert.ToBase64String(fileContent);

            Entity _annotation = new Entity("annotation");
            _annotation.Attributes["objectid"] = new EntityReference("contact", new Guid("9E18BE8E-B078-E711-80FC-70106FAAEAD1"));
            _annotation.Attributes["objecttypecode"] = "contact";
            _annotation.Attributes["subject"] = "Demo";
            _annotation.Attributes["documentbody"] = encodedData;
            _annotation.Attributes["mimetype"] = @"text/plain";
            _annotation.Attributes["notetext"] = "My Sample attachment";
            _annotation.Attributes["filename"] = "sample.txt";
            
            
            service.Create(_annotation);
        }
        static void Main(string[] args)
        {
            Guid entityId = Guid.Parse("9E18BE8E-B078-E711-80FC-70106FAAEAD1");

            NoteAttachment n = new NoteAttachment();
            n.Annotation();
        }
    }
}
