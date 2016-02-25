// <copyright file="PreLeadCreate.cs" company="">
// Copyright (c) 2015 All Rights Reserved
// </copyright>
// <author>GMC</author>
// <date>11/1/2015 12:23:59 PM</date>
// <summary>Implements the PreLeadCreate Process to check for multiple conditions (customer check, duplicate check, retention check, etc.) prior to determine lead assignment.</summary>

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

namespace FP.Pearl.CRM.Plugins
{
    public class PreLeadCreate : IPlugin
    {
        // Create EntityCollection variables to hold results from querying
        EntityCollection resultsDupCust = new EntityCollection();
        EntityCollection accResultsDupCust = new EntityCollection();
        EntityCollection resultsDupLead = new EntityCollection();
        EntityCollection accResultsDupLead = new EntityCollection();
        DataCollection<Entity> entityCollection1 = null;
        DataCollection<Entity> entityCollection2 = null;
        DataCollection<Entity> entityCollection3 = null;

        // Create a Boolean variable to hold result of duplicate customer check
        bool bolDupCustCheck = false;

        // Create a Boolean variable to hold result of duplicate lead check
        bool bolDupLeadCheck = false;

        // Create following variables for assigning and querying process: FirstName, LastName, CompanyName, EmailAddress, Address Number, Zip Code
        string strFirstName = null;
        string strLastName = null;
        string strCompanyName = null;
        string strEmailAddress = null;
        string strAddressLine1 = null;
        string strAddressNumber = null;
        string strZipCode = null;

        // Useful variables
        string strChk4DupCust = null;
        double dDateDiff = 0.0;
        double dDateDiffCompare = 0.0;
        TimeSpan t;
        DateTime dToday = DateTime.Now; 
        DateTime dDaysToExpire;

        public void Execute(IServiceProvider serviceProvider)
        {
            //Extract the tracing service for use in debugging sandboxed plug-ins.
            ITracingService tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));

            // Obtain the execution context from the service provider.
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));

            // The InputParameters collection contains all the data passed in the message request.
            if (context.InputParameters.Contains("Target") && context.InputParameters["Target"] is Entity)
            {
                // Obtain the target entity from the input parameters.
                // Lead Entity coming in
                Entity entity = (Entity)context.InputParameters["Target"];

                // Instantiate Organization Service Interfaces
                IOrganizationServiceFactory servicefactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
                IOrganizationService service = servicefactory.CreateOrganizationService(context.UserId);

                try
                {
                    #region STEP -2 Duplicate Customer Check

                    /* Determine duplication based on following criteria:

                    Customer Check by Search Fields: 
                    o	First Name
                    o	Last Name
                    o	Company Name (Strip abbreviations: CO, INC, THE, OF and CORP)
                    o	Email Address
                    o	Address Number (123 Main St. Use Only 123)
                    o	Zip Code

                    In the rules below, active customer match is defined below:

                    •	Email Address is an exact match then there is a customer match OR
                    •	First Name, Last name, Address Number and Zip Code is a customer match OR
                    •	First Name, Last Name, Zip Code and Company Name is a customer match OR
                    •	Company Name, Address Number and Zip Code is a customer match OR 

                     If a match is found for a Cancelled Customer (An account with no active meter line for more than 6 months), this is NOT a customer match. If an existing customer is a match, check box as “existing customer”.  
                    
                     */

                    // Assign Raw Lead Object elements to created variables
                    strFirstName = getFieldData(entity, "firstname");
                    strLastName = getFieldData(entity, "lastname");
                    strCompanyName = getFieldData(entity, "companyname");
                    strEmailAddress = getFieldData(entity, "emailaddress1");
                    strAddressLine1 = getFieldData(entity, "address1_line1");
                    strZipCode = getFieldData(entity, "address1_postalcode");
                   
                    int pc = -1;
                    
                    if (entity.Attributes.Contains("new_productcategory"))
                    {
                        pc = ((Microsoft.Xrm.Sdk.OptionSetValue)(entity.Attributes["new_productcategory"])).Value;
                    }

                    // Create an object that will contain a list of company abbreviations. 
                    // Use the pearl_anytwokey table in action = 3
                    PearlAnyTwoKeyValue pATK3 = new PearlAnyTwoKeyValue();
                    DataCollection<Entity> entityCollection = pATK3.Execute(serviceProvider, 3);

                    List<string> listCompanyAbbreviations = new List<string>();

                    foreach (pearl_anytwokey plD in entityCollection)
                    {
                        listCompanyAbbreviations.Add(plD.Attributes["pearl_text1"].ToString());
                    }

                    // Strip Company Name abbreviations
                    
                    // Loop through List with foreach.
                    foreach (String strAbbreviation in listCompanyAbbreviations)
                    {
                        strCompanyName = strCompanyName.SafeReplace(strAbbreviation, "", true);
                    }

                    // Strip Address Number from Addresses
                    strAddressNumber = StringExtensions.StripAddressNumber(strAddressLine1);

                    // Strip last 4 numbers from Zip Code
                    strZipCode = StringExtensions.StripZipCode(strZipCode);

                    // Check existing Organization Service alongside the Raw Lead Object for:
                   
                    // Email form Organization Service == Email from Raw Lead Object
                    // (if equal then set Boolean variable to true, otherwise no) 
                    if (DuplicateCustomerCheck(service, "account", new string[] { "emailaddress1" }, new string[] { strEmailAddress }).Entities.Count > 0 && bolDupCustCheck == false)
                    {
                        bolDupCustCheck = true;
                        accResultsDupCust = resultsDupCust;
                        strChk4DupCust = "account";
                    }                   
                        
                    // FirstName, LastName, Address Number, Zip Code from Organization Service == FirstName, LastName, Address Number, Zip Code from Raw Lead Object
                    // (if equal then set Boolean variable to true, otherwise no)
                    else if (DuplicateCustomerCheck(service, "account", new string[] { "firstname", "lastname", "address1_line1", "address1_postalcode" }, new string[] { strFirstName, strLastName, strAddressNumber, strZipCode }).Entities.Count > 0 && bolDupCustCheck == false)
                    {
                        bolDupCustCheck = true;
                        accResultsDupCust = resultsDupCust;
                        strChk4DupCust = "account";
                    }                   

                    // FirstName, LastName, ZipCode, Company Name from Organization Service == FirstName, LastName, ZipCode, Company Name from Raw Lead Object
                    // (if equal then set Boolean variable to true, otherwise no)
                    else if (DuplicateCustomerCheck(service, "account", new string[] { "firstname", "lastname", "address1_postalcode", "companyname" }, new string[] { strFirstName, strLastName, strZipCode, strCompanyName }).Entities.Count > 0 && bolDupCustCheck == false)
                    {
                        bolDupCustCheck = true;
                        accResultsDupCust = resultsDupCust;
                        strChk4DupCust = "account";
                    }                   

                    // Company Name, Address Number, Zip Code from Organization Service == Company Name, Address Number, Zip Code from Raw Lead Object
                    // (if equal then set Boolean variable to true, otherwise no)
                    else if (DuplicateCustomerCheck(service, "account", new string[] { "companyname", "address1_line1", "address1_postalcode" }, new string[] { strCompanyName, strAddressNumber, strZipCode }).Entities.Count > 0 && bolDupCustCheck == false)
                    {
                        bolDupCustCheck = true;
                        accResultsDupCust = resultsDupCust;
                        strChk4DupCust = "account";
                    }

                    // On a positive match process:
                    if (bolDupCustCheck && strChk4DupCust == "account")
                    {
                        // Active Rental Header -- Rental Line -- Deal code that doesn't start with D or C and Has GEN PROD POS = Meter, Document and Line No.
                        // More that one Active Real Header for the same account - BEWARE!!!!!!
                        RentalLineByAccount rLBA = new RentalLineByAccount();
                        entityCollection3 = rLBA.Execute(serviceProvider, strCompanyName);

                        if (entityCollection3.Count > 0)
                        {
                            /* Qualify the Incoming Lead */
                            entity.Attributes["statecode"] = new OptionSetValue(1); // Qualified
                            entity.Attributes["statuscode"] = "Qualified because Active Rental Header";
                        }
                        else
                        {
                            if (accResultsDupCust[0].Attributes.Contains("accountid"))
                            {
                                context.SharedVariables.Add("accountId", accResultsDupCust[0].Attributes["accountid"].ToString());
                            }

                            Guid accId = new Guid(accResultsDupCust[0].Attributes["accountid"].ToString());
                            entity.Attributes["pearl_existingcustomer"] = bolDupCustCheck;
                            entity.Attributes["parentaccountid"] = new EntityReference("account", accId);
                        }
                    }

                    // Nullify some variables for reuse or GC
                    resultsDupCust = null;
                    accResultsDupCust = null;
                    listCompanyAbbreviations = null;
                    bolDupCustCheck = false;
                    strChk4DupCust = null;

                    #endregion

                    #region STEP 3 – Duplicate Check and Retention Check

                    /* Process duplicate check based on following criteria:
                
                    All SELLFP.COM leads are NOT Duplicates. Create new record for them and they will go SellFP email distribution.
                    Check for any duplicate leads. If the lead is NOT a duplicate, it will move to Step 4: Qualify stage.
                    CRM to search duplicate leads by the same Customer Check Rules above except Category Match will be a new field to check. 

                    o	First Name
                    o	Last Name
                    o	Company Name (Strip abbreviations: CO, INC, THE, OF and CORP)
                    o	Email Address
                    o	Address Number (123 Main St. Use Only 123)
                    o	Zip Code
                    o	Category Match (meter, folder/inserter, etc)

                    In the rules below, active duplicate match is defined below:

                    •	If an exact match lead comes in but with two different Categories, it is NOT a duplicate.
                    •	Email Address is an exact match then there is a duplicate match.s
                    •	First Name, Last name, Address Number and Zip Code is a duplicate match.
                    •	First Name, Last Name, Zip Code and Company Name is a duplicate match
                    •	Company Name, Address Number and Zip Code is a duplicate match.

                    Any lead marked as duplicate will link to active lead for reference. 
                    */

                    // Check existing Organization Service alongside the Lead Object and check for:

                    // Email form Organization Service == Email from Lead Object
                    // (if equal then check if Category from Organization Service == Category from Lead Object, if equal then set Boolean variable to true, otherwise no) 
                    if (DuplicateLeadCheck(service, "lead", new string[] { "emailaddress1" }, new string[] { strEmailAddress }, pc).Entities.Count > 0 && bolDupLeadCheck==false)
                    {
                        bolDupLeadCheck = true;
                        accResultsDupLead = resultsDupLead;
                    }

                    // FirstName, LastName, Address Number, Zip Code from Organization Service == FirstName, LastName, Address Number, Zip Code from Lead Object
                    // (if equal then check if  Category from Organization Service == Category from Lead Object, if equal then set Boolean variable to true, otherwise no)
                    else if (DuplicateLeadCheck(service, "lead", new string[] { "firstname", "lastname", "address1_line1", "address1_postalcode" }, new string[] { strFirstName, strLastName, strAddressNumber, strZipCode }, pc).Entities.Count > 0 && bolDupLeadCheck == false)
                    {
                        bolDupLeadCheck = true;
                        accResultsDupLead = resultsDupLead;
                    }

                    // FirstName, LastName, ZipCode, Company Name from Organization Service == FirstName, LastName, ZipCode, Company Name from Lead Object
                    // (if equal then check if Category from Organization Service == Category from Lead Object, if equal then set Boolean variable to true, otherwise no)
                    else if (DuplicateLeadCheck(service, "lead", new string[] { "firstname", "lastname", "companyname", "address1_postalcode" }, new string[] { strFirstName, strLastName, strCompanyName, strZipCode }, pc).Entities.Count > 0 && bolDupLeadCheck == false)
                    {
                        bolDupLeadCheck = true;
                        accResultsDupLead = resultsDupLead;
                    }

                    // Company Name, Address Number, Zip Code from Organization Service == Company Name, Address Number, Zip Code from Lead Object
                    // (if equal then check if Category from Organization Service == Category from Lead Object, if equal then set Boolean variable to true, otherwise no)
                    else if (DuplicateLeadCheck(service, "lead", new string[] { "companyname", "address1_line1", "address1_postalcode" }, new string[] { strCompanyName, strAddressNumber, strZipCode }, pc).Entities.Count > 0 && bolDupLeadCheck == false)
                    {
                        bolDupLeadCheck = true;
                        accResultsDupLead = resultsDupLead;
                    }

                    /* Process retention check based on following criteria:
                
                    If incoming lead has retention check, then disable incoming lead. End of Workflow Process.
                    If original lead has the retention check, then disable the original lead 

                    */

                    // Check if Raw Object Lead is of Retention Check == TRUE, if equal set the Raw Object Lead as disabled, otherwise,
                    // check if Organization Service Lead is of Retention Check == TRUE, if equal set the Organization Service Lead as disabled.
                    if (accResultsDupLead.Entities.Count > 0)
                    {
                        foreach (Entity item in accResultsDupLead.Entities)
                        {
                            if (entity.Attributes.Contains("pearl_retention"))
                            {
                                /* Incoming lead no retention */
                                if (!Convert.ToBoolean(entity.Attributes["pearl_retention"]))
                                {
                                    /* Checking Original Lead for retention Check */
                                    if (item.Attributes["pearl_retention"] != null)
                                    {
                                        if (Convert.ToBoolean(item.Attributes["pearl_retention"]))
                                        {
                                            // Pass the data to the post event plug-in in an execution context shared variable named orginalLeadID
                                            context.SharedVariables.Add("orginalLeadID", item.Attributes["leadid"].ToString());
                                        }
                                        else
                                        {
                                            // pearl_daytoexpire > today ==> disable incoming Lead keep original
                                            dDaysToExpire = Convert.ToDateTime(entity.Attributes["pearl_Daystoexpire"]);

                                            if (dDaysToExpire > dToday)
                                            {
                                                /* Disable Incoming Lead */
                                                entity.Attributes["statecode"] = new OptionSetValue(2); // Disqualified
                                                entity.Attributes["statuscode"] = "Disqualified - DaysToExpire greater than Today";
                                                entity.Attributes["pearl_leadstatus"] = "Disqualified";
                                            }
                                        }
                                    }
                                }
                                else
                                {
                                    entity.Attributes["pearl_linktooriginallead"] = new EntityReference("lead", (Guid)item.Attributes["leadid"]);

                                    /* Disable Incoming Lead */
                                    entity.Attributes["statecode"] = new OptionSetValue(2); // Disqualified
                                    entity.Attributes["statuscode"] = "Disqualified - Lead No Retention";
                                    entity.Attributes["pearl_leadstatus"] = "Disqualified";
                                }
                            }
                        }
                    }

                    // Nullify some variables for reuse or GC
                    resultsDupLead = null;
                    accResultsDupLead = null;
                    bolDupLeadCheck = false;

                    #endregion

                    #region STEP 4 – Qualified

                    /* Qualification process based on the following criteria:
                
                    If pearl_AutoQualified equal true qualify the incoming Lead
                    Else
                    Assign Round Robin to a DDR
                    
                    */

                    // Check for pearl_AutoQualified
                    if ((bool)(entity.Attributes["pearl_AutoQualified"]))
                    {
                        /* Qualify the Incoming Lead */
                        entity.Attributes["statecode"] = new OptionSetValue(1); // Qualified
                        entity.Attributes["statuscode"] = "Qualified because pearl_AutoQualified is true";
                    }
                    else
                    {
                        // Query Expression to get from Lead Routing Type the Default Distribution Account based on pearl_leadRoutingType
                        LeadRoutingTypeDefaultAccount lRTDA = new LeadRoutingTypeDefaultAccount();
                        entityCollection1 = lRTDA.Execute(serviceProvider, entity.Attributes["pearl_leadroutingtype"].ToString());

                        string sDefaultAccount = entityCollection1[0].Attributes["pearl_DefaultAccountName"].ToString();

                        // Query Expression to get from Account all Contacts with Default Distribution Account and "pearl_AssignLead" == true,
                        AllContactsDefaultAccount aCDA = new AllContactsDefaultAccount();
                        entityCollection2 = aCDA.Execute(serviceProvider, sDefaultAccount);

                        string sContactID = "";

                        // Round Robin
                        foreach (Contact cT in entityCollection2)
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
                            ColumnSet = new ColumnSet("pearl_LastOpportunityAssignment", "pearl_AssignLead", "adx_systemuserid")
                        };

                        queryContact1.AddAttributeValue("parentcustomerid", sContactID);
                        EntityCollection entResultContact1 = service.RetrieveMultiple(queryContact1);

                        if (checkLead(entResultContact1))
                        {
                            entity.Attributes["OwnerId"] = entResultContact1[0].Attributes["adx_systemuserid"]; 
                            entResultContact1[0].Attributes["pearl_AssignLead"] = true;
                            entResultContact1[0].Attributes["pearl_LastOpportunityAssignment"] = dToday;
                        }
                    }

                    #endregion

                    #region STEP 5 – Pre Map Check

                    /* Process Pre Map Check based on following criteria:

                    Leads need to be checked if they are a MAP lead, opportunity or existing customer. See separate “MAP Accounts” spreadsheet.

                    MAP check will also include one or all fields:
                    o	Emails ending with: .edu, .gov, .mil
                    o	Emails ending with FP’s existing MAP accounts (i.e. @Walgreens.com, @Lowes.com, etc.)
                    o	MAP Customer Names (The UPS Store, Hilton, etc.)
                
                    If any criteria above is met, check the MAP tracking field.
                
                    */

                    // Create an object that will contain a list of identifying MAP emails fields like: .edu, .gov, .mil, existing MAP accounts, MAP Customer Names, etc. - TO DO Based on CRM
                    // Use the pearl_anytwokey table in action = 1
                    PearlAnyTwoKeyValue pATK1 = new PearlAnyTwoKeyValue();
                    entityCollection = pATK1.Execute(serviceProvider, 1);
                    
                    List<String> listMAPEmails = new List<String>();

                    foreach (pearl_anytwokey plD in entityCollection)
                    {
                        listMAPEmails.Add(plD.Attributes["pearl_text1"].ToString());
                    }

                    // Create a process that will take the Raw Lead Object and check against the object containing a list of identifying MAP email fields,
                    // if found mark the MAP tracking field.

                    // Loop through List with foreach.
                    foreach (String strMAPEmail in listMAPEmails)
                    {
                        // Obtain Raw Lead Object
                        if (entity.Attributes.Contains("emailaddress1"))
                        {
                            string emailString = entity.Attributes["emailaddress1"].ToString().ToLower();
                            if (emailString.Contains(strMAPEmail))
                            {
                                entity.Attributes["pearl_map"] = true;
                            }
                        }
                    }

                    // Create an object that will contain a list of identifying MAP names fields like existing MAP accounts, MAP Customer Names, etc. - TO DO Based on CRM
                    // Use the pearl_anytwokey table in action = 2
                    PearlAnyTwoKeyValue pATK2 = new PearlAnyTwoKeyValue();
                    entityCollection = pATK2.Execute(serviceProvider, 2);

                    List<string> listMAPNames = new List<string>();

                    foreach (pearl_anytwokey plD in entityCollection)
                    {
                        listMAPNames.Add(plD.Attributes["pearl_text1"].ToString());
                    }                    

                    // Create a proess that will take the Raw Lead Object and check against the object containing a list of identifying MAP names,
                    // if found mark the MAP tracking field.

                    // Loop through List with foreach.
                    foreach (String strMAPName in listMAPNames)
                    {
                        // Obtain Raw Lead Object
                        if (entity.Attributes.Contains("company"))
                        {
                            string emailString = entity.Attributes["company"].ToString().ToLower();
                            if (emailString.Contains(strMAPName))
                            {
                                entity.Attributes["pearl_map"] = true;
                            }
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

        /*
        Create a Function that will take the existing Organization Service alongside the Raw Lead Object and check for:
        -	Email form Organization Service == Email from Raw Lead Object
        (if equal then set Boolean variable to true, otherwise no) 
        -	FirstName, LastName, Address Number, Zip Code from Organization Service == FirstName, LastName, Address Number, Zip Code from Raw Lead Object
        (if equal then set Boolean variable to true, otherwise no)
        -	FirstName, LastName, ZipCode, Company Name from Organization Service == FirstName, LastName, ZipCode, Company Name from Raw Lead Object
        (if equal then set Boolean variable to true, otherwise no)
        -	Company Name, Address Number, Zip Code from Organization Service == Company Name, Address Number, Zip Code from Raw Lead Object
        (if equal then set Boolean variable to true, otherwise no)
        */
        public EntityCollection DuplicateCustomerCheck
            (
            IOrganizationService service,
            string checkFor,
            string[] filterFields,
            string[] fieldValues,
            int productCategory = -1
            )
        {
            int i = 0;
            ColumnSet cs = new ColumnSet();
            QueryExpression QE = new QueryExpression(checkFor);
            cs = new ColumnSet("name", "accountid", "statuscode", "statecode");
            QE.ColumnSet = cs;

            foreach (string item in filterFields)
            {
                if (item.Equals("address1_line1"))
                {
                    QE.Criteria.AddCondition(item, ConditionOperator.BeginsWith, fieldValues[i]);
                }
                else if (item.Equals("address1_postalcode"))
                {
                    QE.Criteria.AddCondition(item, ConditionOperator.BeginsWith, fieldValues[i]);
                }
                else if (item.Equals("companyname"))
                {
                    QE.Criteria.AddCondition(item.ToLower(), ConditionOperator.Equal, fieldValues[i].ToLower());
                }
                else if (item.Equals("name"))
                {
                    QE.Criteria.AddCondition(item.ToLower(), ConditionOperator.Equal, fieldValues[i].ToLower());
                }
                else
                {
                    QE.Criteria.AddCondition(item.ToLower(), ConditionOperator.Equal, fieldValues[i].ToLower());
                }
                i++;
            }

            resultsDupCust = service.RetrieveMultiple(QE);
            return resultsDupCust;
        }

        /*
        Create a Function that will take the existing Organization Service alongside the Lead Object and check for:
        -	Email form Organization Service == Email from Lead Object
        (if equal then check if Category from Organization Service == Category from Lead Object, if equal then set Boolean variable to true, otherwise no) 
        -	FirstName, LastName, Address Number, Zip Code from Organization Service == FirstName, LastName, Address Number, Zip Code from Lead Object
        (if equal then check if  Category from Organization Service == Category from Lead Object, if equal then set Boolean variable to true, otherwise no)
        -	FirstName, LastName, ZipCode, Company Name from Organization Service == FirstName, LastName, ZipCode, Company Name from Lead Object
        (if equal then check if Category from Organization Service == Category from Lead Object, if equal then set Boolean variable to true, otherwise no)
        -	Company Name, Address Number, Zip Code from Organization Service == Company Name, Address Number, Zip Code from Lead Object
        (if equal then check if Category from Organization Service == Category from Lead Object, if equal then set Boolean variable to true, otherwise no)
        */
        public EntityCollection DuplicateLeadCheck
            (
            IOrganizationService service,
            string checkFor,
            string[] filterFields,
            string[] fieldValues,
            int productCategory = -1
            )
        {
            int i = 0;
            ColumnSet cs = new ColumnSet();
            QueryExpression QE = new QueryExpression(checkFor);
            cs = new ColumnSet("companyname", "pearl_retention", "statuscode", "pearl_daystoexpire", "parentaccountid", "leadid");
            QE.ColumnSet = cs;

            foreach (string item in filterFields)
            {
                if (item.Equals("address1_line1"))
                {
                    QE.Criteria.AddCondition(item, ConditionOperator.BeginsWith, fieldValues[i]);
                }
                else if (item.Equals("address1_postalcode"))
                {
                    QE.Criteria.AddCondition(item, ConditionOperator.BeginsWith, fieldValues[i]);
                }
                else if (item.Equals("companyname"))
                {
                    QE.Criteria.AddCondition(item.ToLower(), ConditionOperator.Equal, fieldValues[i].ToLower());
                }
                else if (item.Equals("name"))
                {
                    QE.Criteria.AddCondition(item.ToLower(), ConditionOperator.Equal, fieldValues[i].ToLower());
                }
                else
                {
                    QE.Criteria.AddCondition(item.ToLower(), ConditionOperator.Equal, fieldValues[i].ToLower());
                }
                i++;
            }

            if (productCategory != -1)
            {
                QE.Criteria.AddCondition("new_productcategory", ConditionOperator.Equal, productCategory);
            }

            QE.Criteria.AddCondition("statecode", ConditionOperator.LessEqual, 1);  // Look in to open and Qualified.

            resultsDupLead = service.RetrieveMultiple(QE);
            return resultsDupLead;
        }

        // Assign Entity to FieldNames
        public string getFieldData(Entity entity, string fieldName)
        {
            if (entity.Attributes.Contains(fieldName))
            {
                return entity.Attributes[fieldName].ToString();
            }
            else
            {
                return "";
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
    }

    public static class StringExtensions
    {
        // Create a Function to Strip Company Name abbreviations
        public static string SafeReplace(this string input, string find, string replace, bool matchWholeWord)
        {
            string textToFind = matchWholeWord ? string.Format(@"\b{0}\b", find) : find;
            return Regex.Replace(input, textToFind, replace);
        }

        // Create a Function to Strip Address Number from Addresses
        public static string StripAddressNumber(this string input)
        {
             string[] strTemp = input.Split(' ');
             string strResult = null;
             if (strTemp.Length > 0)
             {
                 strResult = strTemp[0];
             }
             return strResult;
        }

        // Create a Function to Strip last 4 numbers from Zip Code
        public static string StripZipCode(this string input)
        {
            string zipcode = input;
            if (zipcode.Length > 4)
                zipcode = zipcode.Substring(0, 5);
            return zipcode;
        }
    }
}
