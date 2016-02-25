// <copyright file="PostOpportunityCreate.cs" company="">
// Copyright (c) 2015 All Rights Reserved
// </copyright>
// <author>GMC</author>
// <date>11/1/2015 12:23:59 PM</date>
// <summary>Implements the PostOpportunityCreate Process to fulfill extra requirements for opportunity objects after assignmeent.</summary>

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Crm.Sdk.Messages;
using System.Text.RegularExpressions;

namespace FP.Pearl.CRM.Plugins
{
    public class PostOpportunityCreate: IPlugin
    {
        EntityCollection results = new EntityCollection();
        EntityCollection accResults = new EntityCollection();
        DateTime dToday = DateTime.Now;

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
                // Opportunity Entity coming in
                Entity entity = (Entity)context.InputParameters["Target"];

                IOrganizationServiceFactory servicefactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
                IOrganizationService service = servicefactory.CreateOrganizationService(context.UserId);

                // Need to map Lead to Opportunity
                PearlAnyTwoKeyMapping pATKM = new PearlAnyTwoKeyMapping();
                DataCollection<Entity>  entityCollection = pATKM.Execute(serviceProvider, 4, "Lead", "Opportunity"); 

                foreach (pearl_anytwokey plD in entityCollection)
                {
                    entity.Attributes[plD.Attributes["pearl_text4"].ToString()] = plD.Attributes["pearl_text2"].ToString();
                }

                // Obtain pearl_daystoexpire from LeadRoutingType
                QueryByAttribute queryLeadRoutingType = new QueryByAttribute
                {
                    EntityName = "LeadRoutingType",
                    ColumnSet = new ColumnSet("pearl_daystoexpire")
                };

                queryLeadRoutingType.AddAttributeValue("pearl_leadroutingtype", entity.Attributes["pearl_routingtype"].ToString());
                EntityCollection entResultAccount = service.RetrieveMultiple(queryLeadRoutingType);

                if (checkLead(entResultAccount))
                {
                    int iDaysToExpire = Int32.Parse(entResultAccount[0].Attributes["pearl_daystoexpire"].ToString());
                    TimeSpan duration = new System.TimeSpan(iDaysToExpire, 0, 0, 0);
                    DateTime answer = dToday.Add(duration);
                    entity.Attributes["pearl_datetoexpire"] = answer;
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
    }
}
