// <copyright file="PostLeadCreate.cs" company="">
// Copyright (c) 2015 All Rights Reserved
// </copyright>
// <author>GMC</author>
// <date>11/1/2015 12:23:59 PM</date>
// <summary>Implements the PostLeadCreate Process to fulfill extra requirements after assignment but prior to opportunity state.</summary>

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.ServiceModel;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Crm.Sdk.Messages;
using System.Text.RegularExpressions;
using Microsoft.Xrm.Sdk.Client;
using System.ServiceModel.Description;

namespace FP.Pearl.CRM.Plugins
{
    public class PostLeadCreate : IPlugin
    {
        string strLeadSourceName = null;
        string sAccountName = null;
        string sAccountId = null;
        string sDefaultContact = null;
        double dDateDiff = 0.0;
        double dDateDiffCompare = 0.0;
        TimeSpan t;
        DateTime dToday = DateTime.Now;
        private ITracingService tracingService;

        public void Execute(IServiceProvider serviceProvider)
        {
            //Extract the tracing service for use in debugging sandboxed plug-ins.
            tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));

            // Obtain the execution context from the service provider.
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));

            // The InputParameters collection contains all the data passed in the message request.
            if (context.InputParameters.Contains("Target") && context.InputParameters["Target"] is Entity)
            {
                // Obtain the target entity from the input parameters.
                // Opportunity Entity coming in
                Entity entity = (Entity)context.InputParameters["Target"];

                IOrganizationServiceFactory servicefactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
                IOrganizationService service = servicefactory.CreateOrganizationService(context.UserId);

                try
                {
                    #region STEP 6 - Business Rules & Assign Sales Channel

                    /* Lead – Sales Channel Assignment, process based on following criteria:

                    Identifies which product type (Seg A, Seg B, F/I and OEM) the lead is interested in and which Sales Channel it gets assigned to.

                    MAP? Yes – Go to Step 7. This includes all Servicing Dealers that are MAP accounts. 
                    Retention? Yes – Go to Step 7. If no, it will go to existing customer check.
                    Cancellation? Check if existing customer? Yes – It will go to the Servicing Dealer. If Servicing Dealer is House (Between 4000 and 4999), it will go to Cancellations Team. If not, it will go to the Dealer.  

                    */

                    // Check if Object Lead is MAP == TRUE, if equal then ASSIGN LEAD, otherwise
                    // Check if Object Lead is Retention == TRUE, if equal then ASSIGN LEAD, otherwise
                    // Check if Object Lead is Cancellation == TRUE, if equal then LEAD ASSIGNED TO CANCELLATIONS TEAM, otherwise
                    // Check if Object Lead is Existing Customer == TRUE, if equal then Check if Servicing Dealer Criteria is met == TRUE, if equal then LEAD ASSIGNED TO CANCELLATIONS TEAM, otherwise LEAD ASSIGNED TO DEALER.

                    // Obtain the Lead Entity Object based on Opportunity Entity Originating Lead Id
                    QueryByAttribute queryLeadEntity = new QueryByAttribute
                    {
                        EntityName = "Lead",
                        ColumnSet = new ColumnSet("pearl_saleschannel")
                    };

                    queryLeadEntity.AddAttributeValue("leadid", entity.Attributes["originatingleadid"]);
                    EntityCollection entResultLeadEntity = service.RetrieveMultiple(queryLeadEntity);

                    if (checkLeadPos(entResultLeadEntity))
                    {
                         /*  Check for MAP */
                        if(entResultLeadEntity[0].Attributes["pearl_saleschannel"].Equals("MAP"))
                        {
                            entity.Attributes["pearl_leadroutingtype"] = "MAP";
                        }
                        /*  Check for Retention */
                        else if(entResultLeadEntity[0].Attributes["pearl_saleschannel"].Equals("Retention"))
                        {
                            entity.Attributes["pearl_leadroutingtype"] = "Retention";
                        }
                        /*  Check for Cancellations */
                        else if(entResultLeadEntity[0].Attributes["pearl_saleschannel"].Equals("Cancellations"))
                        {
                            entity.Attributes["pearl_leadroutingtype"] = "Cancellations";
                        }
                        else
                        {
                             /*  Check for EXISTING CUSTOMER */
                            if((bool)entResultLeadEntity[0].Attributes["pearl_existingcustomer"])
                            {
                                /*  Check for SERVICING DEALER CRITERIA */
                                if(entResultLeadEntity[0].Attributes["new_servicingsalesperson"].Equals("")) // TO DO - Check with Alex if this is still valid ????
                                {
                                    entity.Attributes["pearl_leadroutingtype"] = "Cancellations";
                                }
                                else
                                {
                                    entity.Attributes["pearl_leadroutingtype"] = "Dealer";
                                }
                            }
                        }

                        /* New Sales Product Type Distribution, process based on following criteria:

                        To determine product type
                        •	If meter, determine if A or B segment
                        •	If NOT meter, determine if F/I or OEM
                        •	If NO product type is given, A segment is default

                        */

                        // Check if Object Lead is meter == TRUE, if equal then check if Object Lead is segment == A || == B, if equal then assign product type = meter, otherwise check if Object Lead is product type == F/I || == OEM,
                        // if equal then assign product type == to F/I or OEM, otherwise check if Object Lead is product type == “”, if equal then assign product type = meter, segment = A.

                        /*  Check for NEW CURRENT METER */
                        if(entResultLeadEntity[0].Attributes["new_currentmeter"].Equals(""))
                        {
                            /*  Check for SEGMENT NOT A or NOT B */
                            if(!entResultLeadEntity[0].Attributes["pearl_producttype"].Equals("Seg A") || !entResultLeadEntity[0].Attributes["pearl_producttype"].Equals("Seg B"))
                            {
                                entity.Attributes["pearl_leadtype"] = "Seg A";
                            }
                        }
                    }

                    #endregion

                    #region STEP 7 - Assign Lead

                    /* Process based on following criteria: */

                    // NEW PROCESS

                    /*
                    Opportunity to Lead Routing Type to Lead Distribution to get All Lead Distributions that have same Lead Source as the Opportunity
                    Opportunity to Lead Source to get All Lead Source and compare against view above with a Start Date and End Date.
                    From Zip Code to Counties to Account to get All Account by Zip Code and County and compare against view result from above.
                    All accounts available for Distribution.
                    Round Robin in Account and Round Robin in the Contact to Assign the LEAD
                    IF NOTHING IS FOUND:
                    Opportunity to Lead Routing Type to Account to get Default Account Information
                    Round Robin for the Contact associated with Default Contact Above.
                    IF NO CONTACTS AVAILABLE THEN ASSIGN TO DEFAULT CONTACT in the Lead Routing Type.
                    */

                    // Obtain Lead Source Name
                    if (context.InputParameters.Contains("pearl_leadsourcename"))
                    {
                        /*  Check for LEAD SOURCE NAME */
                        QueryByAttribute queryLeadSourceName = new QueryByAttribute
                        {
                            EntityName = "pearl_leadsourcename",
                            ColumnSet = new ColumnSet("pearl_leadsourceid", String.Empty)
                        };

                        queryLeadSourceName.AddAttributeValue("pearl_leadsourceid", ((EntityReference)entity.Attributes["pearl_leadsourcename"]).Id);
                        EntityCollection entResultLeadSourceName = service.RetrieveMultiple(queryLeadSourceName);

                        if (checkLead(entResultLeadSourceName))
                        {
                            strLeadSourceName = entity.Attributes["pearl_leadsourcename"].ToString();
                        }
                    }

                    AllLeadSourceEqualOpportunity aLSEO = new AllLeadSourceEqualOpportunity();
                    DataCollection<Entity>  entityCollection1 = aLSEO.Execute(serviceProvider, strLeadSourceName); 

                    List<string> lNames = new List<string>();

                    foreach (pearl_leaddistribution plD in entityCollection1)
                    {
                        if(plD.pearl_StartDate == null && plD.pearl_EndDate == null)
                        {
                            lNames.Add(plD.pearl_accountname);
                        }
                        else if (plD.pearl_StartDate == null && plD.pearl_EndDate != null)
                        {
                            // (DATE NOW BEFORE END == Good / DATE NOW AFTER ED == Bad)
                            if (dToday < plD.pearl_EndDate)
                            {
                                lNames.Add(plD.pearl_accountname);
                            }
                        }
                        else if(plD.pearl_StartDate != null && plD.pearl_EndDate == null)
                        {
                            // (DATE NOW AFTER SD == GOOD / DATE NOW BEFORE SD == BAD)
                            if (dToday > plD.pearl_StartDate)
                            {
                                lNames.Add(plD.pearl_accountname);
                            }
                        }
                        else
                        {
                            // DEPENDS ON DATE NOW
                            if ((plD.pearl_StartDate < dToday) && (dToday < plD.pearl_EndDate))
                            {
                                lNames.Add(plD.pearl_accountname);
                            }
                        }
                    }

                    List<string> lSource = new List<string>();
                    AllLeadSourceStartDateEndDate aLSSDED = new AllLeadSourceStartDateEndDate();

                    foreach (string lN in lNames)
                    {
                        DataCollection<Entity>  entityCollection2 = aLSSDED.Execute(serviceProvider, lN);

                        foreach (pearl_leadsource plS in entityCollection2)
                        {
                            lSource.Add(plS.pearl_name);
                        }
                    }

                    List<string> lAccount = new List<string>();
                    AllAccountNames aAN = new AllAccountNames();

                    foreach (string lS in lSource)
                    {
                        DataCollection<Entity>  entityCollection3 = aAN.Execute(serviceProvider, lS);

                        foreach (Account aC in entityCollection3)
                        {
                            lAccount.Add(aC.Name);
                        }
                    }

                    DataCollection<Entity> entityCollection4 = null;

                    if (lAccount.Count < 0)
                    {
                        // IF NOTHING IS FOUND:

                        // Opportunity to Lead Routing Type to Account to get Default Account Information
                        List<string> lACountLeadSource = new List<string>();
                        AllAccountLeadSource aALS = new AllAccountLeadSource();
                        entityCollection4 = aALS.Execute(serviceProvider, strLeadSourceName);

                        foreach (Account aCLS in entityCollection4)
                        {
                            lACountLeadSource.Add(aCLS.AccountId.ToString());
                        }

                        if (lACountLeadSource.Count < 0)
                        {
                            // IF NO CONTACTS AVAILABLE THEN ASSIGN TO DEFAULT CONTACT in the Lead Routing Type.
                            
                            // Obtain Default Contact Name
                            LeadDistributionDefaultContact lDDC = new LeadDistributionDefaultContact();
                            DataCollection<Entity>  entityCollection5 = lDDC.Execute(serviceProvider, strLeadSourceName);

                            foreach (pearl_leadroutingtype pRT in entityCollection5)
                            {
                                Contact cDC = pRT.pearl_contact_pearl_leadroutingtype_DefaultContact;
                                sDefaultContact = cDC.ContactId.ToString();
                            }

                            // Assign the Lead to the Contact
                            QueryByAttribute queryContact3 = new QueryByAttribute
                            {
                                EntityName = "Contact",
                                ColumnSet = new ColumnSet("pearl_LastOpportunityAssignment", "pearl_AssignLead")
                            };

                            queryContact3.AddAttributeValue("parentcustomerid", sDefaultContact);
                            EntityCollection entResultContact3 = service.RetrieveMultiple(queryContact3);

                            if (checkLead(entResultContact3))
                            {
                                if (entResultContact3[0].Attributes["adx_systemuserid"] != null)
                                {
                                    entity.Attributes["owninguser"] = entResultContact3[0].Attributes["adx_systemuserid"];
                                }

                                entResultContact3[0].Attributes["pearl_AssignLead"] = true;
                                entResultContact3[0].Attributes["pearl_LastOpportunityAssignment"] = dToday;
                                entity.Attributes["msa_partnerid"] = entResultContact3[0].Attributes["parentcustomerid"];
                                entity.Attributes["msa_partneroppid"] = sDefaultContact;
                                entity.Attributes["pearl_sellingdealer"] = entResultContact3[0].Attributes["parentcustomerid"];
                                entity.Attributes["pearl_servicingdealer"] = entResultContact3[0].Attributes["parentcustomerid"];
                            }
                        }
                        else
                        {
                            // Round Robin for the Contact associated with Default Contact Above.

                            // Obtain the Contacts Data Collection
                            AllContactsLeadAssigned aCLA = new AllContactsLeadAssigned();
                            DataCollection<Entity>  entityCollection6 = aCLA.Execute(serviceProvider, sAccountId);

                            // Round Robin in Contact
                            dDateDiff = 0.0;
                            dDateDiffCompare = 0.0;
                            string sContactID = "";

                            foreach (Contact cT in entityCollection6)
                            {
                                t = dToday - (DateTime)cT.pearl_LastOpportunityAssignment;
                                dDateDiffCompare = t.TotalDays;

                                if (dDateDiffCompare > dDateDiff)
                                {
                                    sContactID = cT.ContactId.ToString();
                                    dDateDiff = dDateDiffCompare;
                                }
                            }

                            // Assign the Lead to the Contact
                            QueryByAttribute queryContact1 = new QueryByAttribute
                            {
                                EntityName = "Contact",
                                ColumnSet = new ColumnSet("pearl_LastOpportunityAssignment", "pearl_AssignLead")
                            };

                            queryContact1.AddAttributeValue("parentcustomerid", sContactID);
                            EntityCollection entResultContact1 = service.RetrieveMultiple(queryContact1);

                            if (checkLead(entResultContact1))
                            {
                                if (entResultContact1[0].Attributes["adx_systemuserid"] != null)
                                {
                                    entity.Attributes["owninguser"] = entResultContact1[0].Attributes["adx_systemuserid"];
                                }

                                entResultContact1[0].Attributes["pearl_AssignLead"] = true;
                                entResultContact1[0].Attributes["pearl_LastOpportunityAssignment"] = dToday;
                                entity.Attributes["msa_partnerid"] = entResultContact1[0].Attributes["parentcustomerid"];
                                entity.Attributes["msa_partneroppid"] = sContactID;
                                entity.Attributes["pearl_sellingdealer"] = entResultContact1[0].Attributes["parentcustomerid"];
                                entity.Attributes["pearl_servicingdealer"] = entResultContact1[0].Attributes["parentcustomerid"];
                            }
                        }
                    }
                    else
                    {
                        // Round Robin in Account
                        foreach (Account aC in entityCollection4)
                        {
                            t = dToday - (DateTime)aC.pearl_LastOpportunityAssignment;
                            dDateDiffCompare = t.TotalDays;

                            if (dDateDiffCompare > dDateDiff)
                            {
                                sAccountName = aC.Name;
                                sAccountId = aC.AccountId.ToString();
                                dDateDiff = dDateDiffCompare;
                            }
                        }

                        // Assign the Lead to the Account
                        QueryByAttribute queryAccount1 = new QueryByAttribute
                        {
                            EntityName = "Account",
                            ColumnSet = new ColumnSet("pearl_LastOpportunityAssignment")
                        };

                        queryAccount1.AddAttributeValue("AccountID", sAccountId);
                        EntityCollection entResultAccount1 = service.RetrieveMultiple(queryAccount1);

                        if (checkLead(entResultAccount1))
                        {
                            entResultAccount1[0].Attributes["pearl_LastOpportunityAssignment"] = dToday;
                        }

                        // Special Business Process

                        // Obtain Lead Distribution Special Processing Rules based on Opportunity (Lead Source and Lead Routing Type)
                        LeadDistributionLSLRT lLDLSLRT = new LeadDistributionLSLRT();
                        DataCollection<Entity>  entityCollection8 = lLDLSLRT.Execute(serviceProvider, entity.Attributes["pearl_leadsource"].ToString(), entity.Attributes["pearl_leadroutingtype"].ToString());

                        if((bool) entityCollection8[0].Attributes["specialprocessingrules"])
                        {
                            // Obtain Lead Routing Type Default Account ID based on Opportunity (Lead Routing Type)
                            LeadRoutingTypeOLRT lRTOLRT = new LeadRoutingTypeOLRT();
                            DataCollection<Entity>  entityCollection9 = lRTOLRT.Execute(serviceProvider, entity.Attributes["pearl_leadroutingtype"].ToString());

                            string sDefaultAccountID = entityCollection9[0].Attributes["defaultaccountid"].ToString();

                            // Assign the Lead to the Account
                            QueryByAttribute queryAccount = new QueryByAttribute
                            {
                                EntityName = "Account",
                                ColumnSet = new ColumnSet("AccountID")
                            };

                            queryAccount.AddAttributeValue("AccountID", sDefaultAccountID);
                            EntityCollection entResultAccount = service.RetrieveMultiple(queryAccount);

                            string sAcId = null;

                            if (checkLead(entResultAccount1))
                            {
                                sAcId = entResultAccount[0].Attributes["AccountID"].ToString(); ;
                            }

                            // Obtain the Contacts Data Collection
                            AllContactsLeadAssigned aCLA = new AllContactsLeadAssigned();
                            DataCollection<Entity> entityCollection5 = aCLA.Execute(serviceProvider, sAccountId);

                            // Round Robin in Contact
                            dDateDiff = 0.0;
                            dDateDiffCompare = 0.0;
                            string sContactID = "";

                            foreach (Contact cT in entityCollection5)
                            {
                                t = dToday - (DateTime)cT.pearl_LastOpportunityAssignment;
                                dDateDiffCompare = t.TotalDays;

                                if (dDateDiffCompare > dDateDiff)
                                {
                                    sContactID = cT.ContactId.ToString();
                                    dDateDiff = dDateDiffCompare;
                                }
                            }

                            // Assign the Lead to the Contact
                            QueryByAttribute queryContact2 = new QueryByAttribute
                            {
                                EntityName = "Contact",
                                ColumnSet = new ColumnSet("pearl_LastOpportunityAssignment", "pearl_AssignLead")
                            };

                            queryContact2.AddAttributeValue("parentcustomerid", sContactID);
                            EntityCollection entResultContact2 = service.RetrieveMultiple(queryContact2);

                            if (checkLead(entResultContact2))
                            {
                                if (entResultContact2[0].Attributes["adx_systemuserid"] != null)
                                {
                                    entity.Attributes["owninguser"] = entResultContact2[0].Attributes["adx_systemuserid"]; 
                                }

                                entResultContact2[0].Attributes["pearl_AssignLead"] = true;
                                entResultContact2[0].Attributes["pearl_LastOpportunityAssignment"] = dToday;
                                entity.Attributes["msa_partnerid"] = sAcId;
                                entity.Attributes["msa_partneroppid"] = sContactID;
                                entity.Attributes["pearl_sellingdealer"] = sAcId;
                                entity.Attributes["pearl_servicingdealer"] = sAccountId;
                            }
                        }
                        else
                        {
                            // Obtain the Contacts Data Collection
                            AllContactsLeadAssigned aCLA = new AllContactsLeadAssigned();
                            DataCollection<Entity> entityCollection5 = aCLA.Execute(serviceProvider, sAccountId);

                            // Round Robin in Contact
                            dDateDiff = 0.0;
                            dDateDiffCompare = 0.0;
                            string sContactID = "";

                            foreach (Contact cT in entityCollection5)
                            {
                                t = dToday - (DateTime)cT.pearl_LastOpportunityAssignment;
                                dDateDiffCompare = t.TotalDays;

                                if (dDateDiffCompare > dDateDiff)
                                {
                                    sContactID = cT.ContactId.ToString();
                                    dDateDiff = dDateDiffCompare;
                                }
                            }

                            // Assign the Lead to the Contact
                            QueryByAttribute queryContact2 = new QueryByAttribute
                            {
                                EntityName = "Contact",
                                ColumnSet = new ColumnSet("pearl_LastOpportunityAssignment", "pearl_AssignLead")
                            };

                            queryContact2.AddAttributeValue("parentcustomerid", sContactID);
                            EntityCollection entResultContact2 = service.RetrieveMultiple(queryContact2);

                            if (checkLead(entResultContact2))
                            {
                                if (entResultContact2[0].Attributes["adx_systemuserid"] != null)
                                {
                                    entity.Attributes["owninguser"] = entResultContact2[0].Attributes["adx_systemuserid"]; 
                                }

                                entResultContact2[0].Attributes["pearl_AssignLead"] = true;
                                entResultContact2[0].Attributes["pearl_LastOpportunityAssignment"] = dToday;
                                entity.Attributes["msa_partnerid"] = entResultContact2[0].Attributes["parentcustomerid"];
                                entity.Attributes["msa_partneroppid"] = sContactID;
                                entity.Attributes["pearl_sellingdealer"] = entResultContact2[0].Attributes["parentcustomerid"];
                                entity.Attributes["pearl_servicingdealer"] = entResultContact2[0].Attributes["parentcustomerid"];
                            }
                        }
                    }

                    /*  
                    Afterwards, the lead will go to the primary contact at the enabled Dealer with a link to the 
                    opportunity information. The Dealer will have to login to view info in Portal. From Portal,
                    they will be able to view full copy of the opportunity and acknowledge they received the lead,
                    edit content, qualify (criteria) and complete order.

                    RSM’s get a full copy of the opportunity via email with a link the opportunity. The RSM has
                    the ability to view info in CRM, edit content and qualify (criteria). 

                                    RSM                         DEALER
                    EMAIL	        Full opportunity Info	    View opportunity link only
                    ABILITIES	    View info in CRM            Edit Content
                                    Qualify (criteria)          View full opportunity info in Portal
                                    Edit Content                Qualify (criteria)
                                                                Complete Order

                    Sales Channel Default Distribution

                    If no email address is on file for sales channel, send Opportunity to corresponding manager below.
                    MAP – emails will go to MAP Manager
                    Retention - emails will go to Cancellations Manager
                    Cancellation - emails will go to Cancellations Manager
                    Inside Sales -- emails will go to Inside Sales Manager
                    Dealers – emails will go to Marketing (marketing@fp-usa.com) 
                    */

                    // TO DO - Email Variables used
                    string sFirstName = null;
                    string sLastName = null;
                    string sEmailAddress = null;
                    string sSubject = null;
                    string sDescription = null;
                    string sOpportunityInformation = null;

                    // TO DO - If no email address is on file for sales channel, send Opportunity to corresponding manager below.
                    // TO DO - Which is the object that contains the email address for sales channel????
                    int i = 0;
                    if(i > 1)
                    {
                        sOpportunityInformation = "URL Link TBD Later"; // TO DO - Obtain information
                        SendEmail(tracingService, sFirstName, sLastName, sEmailAddress, sSubject, sDescription, sOpportunityInformation);
                    }
                    else
                    {
                         /*  Check for MAP */
                        if (entResultLeadEntity[0].Attributes["pearl_saleschannel"].Equals("MAP"))
                        {
                            sFirstName = "MAP Manager FirstName"; // TO DO - Obtain information
                            sLastName = "Fitzpatrick";
                            sEmailAddress = "dfitzpatrick@fp-usa.com";
                            sSubject = null;
                            sDescription = null;
                            sOpportunityInformation = "URL Link TBD Later";

                            SendEmail(tracingService, sFirstName, sLastName, sEmailAddress, sSubject, sDescription, sOpportunityInformation);
                        }
                        else if (entResultLeadEntity[0].Attributes["pearl_saleschannel"].Equals("Retention"))
                        {
                            sFirstName = "Retention Manager FirstName"; // TO DO - Obtain information
                            sLastName = "Hannon";
                            sEmailAddress = "mhannon@fp-usa.com";
                            sSubject = null;
                            sDescription = null;
                            sOpportunityInformation = "URL Link TBD Later";

                            SendEmail(tracingService, sFirstName, sLastName, sEmailAddress, sSubject, sDescription, sOpportunityInformation);
                        }
                        else if (entResultLeadEntity[0].Attributes["pearl_saleschannel"].Equals("Cancellations"))
                        {
                            sFirstName = "Cancellation Manager FirstName"; // TO DO - Obtain information
                            sLastName = "Hannon";
                            sEmailAddress = "mhannon@fp-usa.com";
                            sSubject = null;
                            sDescription = null;
                            sOpportunityInformation = "URL Link TBD Later";

                            SendEmail(tracingService, sFirstName, sLastName, sEmailAddress, sSubject, sDescription, sOpportunityInformation);
                        }
                        else if (entResultLeadEntity[0].Attributes["pearl_saleschannel"].Equals("Inside Sales"))
                        {
                            sFirstName = "Inside Sales Manager FirstName"; // TO DO - Obtain information
                            sLastName = "Charatin";
                            sEmailAddress = "dcharatin@fp-usa.com";
                            sSubject = null;
                            sDescription = null;
                            sOpportunityInformation = "URL Link TBD Later";

                            SendEmail(tracingService, sFirstName, sLastName, sEmailAddress, sSubject, sDescription, sOpportunityInformation);
                        }
                        else if (entResultLeadEntity[0].Attributes["pearl_saleschannel"].Equals("Dealer"))
                        {
                            sFirstName = "Dealers Manager FirstName"; // TO DO - Obtain information
                            sLastName = "Thompson";
                            sEmailAddress = "marketing@fp-usa.com;kthompson@fp-usa.com";
                            sSubject = null;
                            sDescription = null;
                            sOpportunityInformation = "URL Link TBD Later";

                            SendEmail(tracingService, sFirstName, sLastName, sEmailAddress, sSubject, sDescription, sOpportunityInformation);
                        }
                    }

                    #endregion
                }
                catch (FaultException<OrganizationServiceFault> ex)
                {
                    throw new InvalidPluginExecutionException("An error occurred in the InsertLeadProcess plug-in.", ex);
                }
                catch (Exception ex)
                {
                    tracingService.Trace("InsertLeadProcess: {0}", ex.ToString());
                    throw;
                }
            }
        }

        public Boolean checkLead(EntityCollection ent)
        {
            Boolean bolResult = false;
            if (ent.Entities.Count <= 0)
            {
                bolResult = true;
            }
            return bolResult;
        }

        public Boolean checkLeadPos(EntityCollection ent)
        {
            Boolean bolResult = false;
            if (ent.Entities.Count >= 0)
            {
                bolResult = true;
            }
            return bolResult;
        }

        public void SendEmail(ITracingService tracingService, string sFirstName, string sLastName, string sEmailAddress, string sSubject, string sDescription, string sOpportunityInformation)
        {
            try
            {
                // Obtain the target organization's Web address and client logon 
                // credentials from the user.
                ServerConnection serverConnect = new ServerConnection();
                ServerConnection.Configuration config = serverConnect.GetServerConfiguration();

                SendEmail app = new SendEmail();
                app.Run(config, tracingService, sFirstName, sLastName, sEmailAddress, sSubject, sDescription, sOpportunityInformation);

            }
            catch (FaultException<Microsoft.Xrm.Sdk.OrganizationServiceFault> ex)
            {
                tracingService.Trace("InsertLeadProcess: {0}", "The application terminated with an error.");
                tracingService.Trace("Timestamp: {0}", ex.Detail.Timestamp);
                tracingService.Trace("Code: {0}", ex.Detail.ErrorCode);
                tracingService.Trace("Message: {0}", ex.Detail.Message);
                tracingService.Trace("Plugin Trace: {0}", ex.Detail.TraceText);
                tracingService.Trace("Inner Fault: {0}",
                    null == ex.Detail.InnerFault ? "No Inner Fault" : "Has Inner Fault");
            }
            catch (System.TimeoutException ex)
            {
                tracingService.Trace("InsertLeadProcess: {0}", "The application terminated with an error.");
                tracingService.Trace("Message: {0}", ex.Message);
                tracingService.Trace("Stack Trace: {0}", ex.StackTrace);
                tracingService.Trace("Inner Fault: {0}",
                    null == ex.InnerException.Message ? "No Inner Fault" : ex.InnerException.Message);
            }
            catch (System.Exception ex)
            {
                tracingService.Trace("InsertLeadProcess: {0}", "The application terminated with an error.");
                tracingService.Trace(ex.Message);

                // Display the details of the inner exception.
                if (ex.InnerException != null)
                {
                    tracingService.Trace(ex.InnerException.Message);

                    FaultException<Microsoft.Xrm.Sdk.OrganizationServiceFault> fe =
                        ex.InnerException
                        as FaultException<Microsoft.Xrm.Sdk.OrganizationServiceFault>;
                    if (fe != null)
                    {
                        tracingService.Trace("Timestamp: {0}", fe.Detail.Timestamp);
                        tracingService.Trace("Code: {0}", fe.Detail.ErrorCode);
                        tracingService.Trace("Message: {0}", fe.Detail.Message);
                        tracingService.Trace("Plugin Trace: {0}", fe.Detail.TraceText);
                        tracingService.Trace("Inner Fault: {0}",
                            null == fe.Detail.InnerFault ? "No Inner Fault" : "Has Inner Fault");
                    }
                }
            }
            // Additonal exceptions to catch: SecurityTokenValidationException, ExpiredSecurityTokenException,
            // SecurityAccessDeniedException, MessageSecurityException, and SecurityNegotiationException.

            finally
            {
                // Any other code to finalize process if needed.
            }
        }
    }
}