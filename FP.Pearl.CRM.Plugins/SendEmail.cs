using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.ServiceModel;
using System.ServiceModel.Description;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk.Discovery;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Client;
using Microsoft.Crm.Sdk.Messages;

namespace FP.Pearl.CRM.Plugins
{
    public class SendEmail
    {
        // Define the IDs needed.
        private Guid _emailId;
        private Guid _contactId;
        private Guid _userId;
        private OrganizationServiceProxy _serviceProxy;
        private ITracingService _tracingService;

        /// <summary>
        /// Send an e-mail message.
        /// </summary>
        /// <param name="serverConfig">Contains server connection information.</param>
        public void Run(ServerConnection.Configuration serverConfig, ITracingService tracingService, string sFirstName, string sLastName, string sEmailAddress, string sSubject, string sDescription, string sOpportunityInformation)
        {
            try
            {
                // Assign ITracingService for future use
                _tracingService = tracingService;

                // Connect to the Organization service. 
                // The using statement assures that the service proxy will be properly disposed.
                using (_serviceProxy = ServerConnection.GetOrganizationProxy(serverConfig))
                {
                    // This statement is required to enable early-bound type support.
                    _serviceProxy.EnableProxyTypes();

                    // Call the method to create any data.
                    CreateRequiredRecords(sFirstName, sLastName, sEmailAddress, sSubject, sDescription, sOpportunityInformation);

                    // Use the SendEmail message to send an e-mail message.
                    SendEmailRequest sendEmailreq = new SendEmailRequest
                    {
                        EmailId = _emailId,
                        TrackingToken = "",
                        IssueSend = true
                    };

                    SendEmailResponse sendEmailresp = (SendEmailResponse)_serviceProxy.Execute(sendEmailreq);
                    _tracingService.Trace("Sent the e-mail message.");

                    DeleteRequiredRecords();
                }
            }

            // Catch any service fault exceptions that Microsoft Dynamics CRM throws.
            catch (FaultException<Microsoft.Xrm.Sdk.OrganizationServiceFault>)
            {
                // You can handle an exception here or pass it back to the calling method.
                throw;
            }
        }

        /// <summary>
        /// This method creates any entity records.
        /// </summary>
        public void CreateRequiredRecords(string sFirstName, string sLastName, string sEmailAddress, string sSubject, string sDescription, string sOpportunityInformation)
        {
            // Create a contact to send an email to (To: field)
            Contact emailContact = new Contact()
            {
                FirstName = sFirstName,
                LastName = sLastName,
                EMailAddress1 = sEmailAddress
            };
            _contactId = _serviceProxy.Create(emailContact);
            _tracingService.Trace("Created a contact: {0}.", emailContact.FirstName + " " + emailContact.LastName);

            // Get a system user to send the email (From: field)
            WhoAmIRequest systemUserRequest = new WhoAmIRequest();
            WhoAmIResponse systemUserResponse = (WhoAmIResponse)_serviceProxy.Execute(systemUserRequest);
            _userId = systemUserResponse.UserId;

            // Create the 'From:' activity party for the email
            ActivityParty fromParty = new ActivityParty
            {
                PartyId = new EntityReference(SystemUser.EntityLogicalName, _userId)
            };

            // Create the 'To:' activity party for the email
            ActivityParty toParty = new ActivityParty
            {
                // PartyId = new EntityReference(Contact.EntityLogicalName, _contactId)
                PartyId = new EntityReference(Contact.EntityLogicalName, _contactId)
            };
            _tracingService.Trace("Created activity parties.");

            // Create an e-mail message.
            Email email = new Email
            {
                To = new ActivityParty[] { toParty },
                From = new ActivityParty[] { fromParty },
                Subject = sSubject,
                Description = sDescription,
                DirectionCode = true
            };
            _emailId = _serviceProxy.Create(email);
            _tracingService.Trace("Create {0}.", email.Subject);
        }

        /// <summary>
        /// Deletes the custom entity record that was created.
        /// </summary>
        public void DeleteRequiredRecords()
        {
            _serviceProxy.Delete(Email.EntityLogicalName, _emailId);
            _serviceProxy.Delete(Contact.EntityLogicalName, _contactId); ;

            _tracingService.Trace("Entity records have been deleted.");
        }
    }
}
