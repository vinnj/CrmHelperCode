using System;
using System.Linq;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Tooling.Connector;
using Microsoft.Crm.Sdk.Messages;
using System.Configuration;

namespace UpdateCRM
{
    public class Academics
    {
        //private IOrganizationService _orgService;
        public Guid RetrievedContactID { get; set; }
        public string Email { get; set; }
        public string InstitutionID { get; set; }
        public string AfterName { get; set; }
        public string SelectedPublications { get; set; }
        public string Interest { get; set; }
        public string ProExperience { get; set; }
        public string Expertise { get; set; }

        #region RetrieveUpdateaAcademics
        /*METHOD TO RETRIEVE XL DATA AND PASSING IT TO RetrieveUpdate METHOD
        THE METHOD CALLS UPDATEHELPER CLASS WHICH RETRIEVES THE XL COLUMN DATA
        THEN PASSES THE XL DCOLUMN DATA TO RETRIEVEUPDATE*/
        public void RetrieveUpdatedAcademics()
        {
            UpdateHelper.UpdateAcademics academicHelper = new UpdateHelper.UpdateAcademics();
            var academicExcel = academicHelper.ReadAcademicData();

            var academics = from p in academicExcel.Worksheet<UpdateHelper.UpdateAcademics>("Academics")
                            select p;

            foreach (var contact in academics)
            {
                Email = contact.Email;
                RetrievedContactID = new Guid(contact.ContactID);
                ProExperience = contact.ProExperience;
                Interest = contact.Interest;
                SelectedPublications = contact.SelectedPublications;
                InstitutionID = contact.InstitutionID;
                AfterName = contact.AfterName;
                Expertise = contact.Expertise;

                Academic(Email, RetrievedContactID, ProExperience, Interest, SelectedPublications, InstitutionID, AfterName, Expertise);
            }
        }
        #endregion

        #region Academic
        /*METHOD TO CONNECT WITH CRM,RETRIEVE & UPDATE DATA.
        RETRIEVE DATA VIA LINQ QUERY
        PASSING EMAIL/CONTACT ID TO UPDATE CRM RECORD
        */
        private void Academic(string Email, Guid Retrieved, string ProExperience, string Interest, string SelectedPublications,
            string InstitutionID, string AfterName, string Expertise)
        {
            
            string connStr = ConfigurationManager.ConnectionStrings[1].ConnectionString;

            CrmServiceClient conn = new CrmServiceClient(connStr);

            // Cast the proxy client to the IOrganizationService interface.
            //_orgService = (IOrganizationService)conn.OrganizationWebProxyClient != null ? (IOrganizationService)conn.OrganizationWebProxyClient : (IOrganizationService)conn.OrganizationServiceProxy;
            IOrganizationService service = (IOrganizationService)conn.OrganizationServiceProxy;
           
            BUCKSdev svcContext = new BUCKSdev(service);

            var contacts = from c in svcContext.ContactSet
                           where c.ContactId == RetrievedContactID
                               select c;

                foreach (var contact in contacts)
                {
                    Contact updateContact = new Contact
                    {
                        ContactId = RetrievedContactID,
                        EMailAddress1 = Email,
                        bnu_ScholarlySummary = Interest,
                        bnu_ProfessionalExperience = ProExperience,
                        bnu_SelectedPublications = SelectedPublications,
                        bnu_Lettersaftername = AfterName,
                        bnu_Bio = Expertise
                             
                        //_serviceProxy.Update(updateContact);

                    };

                Console.WriteLine("Updating {0} ", contact.FullName);
                service.Update(updateContact);
               
            }
        }
        #endregion
    }
}
