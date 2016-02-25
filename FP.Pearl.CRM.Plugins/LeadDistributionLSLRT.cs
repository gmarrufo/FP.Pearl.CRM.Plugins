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
    public class LeadDistributionLSLRT
    {
        EntityCollection results = new EntityCollection();
        EntityCollection accResults = new EntityCollection();
        DataCollection<Entity> entityCollection = null;

        public DataCollection<Entity> Execute(IServiceProvider serviceProvider, string sLeadSource, string sLeadRoutingType)
        {
            //Extract the tracing service for use in debugging sandboxed plug-ins.
            ITracingService tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));

            // Obtain the execution context from the service provider.
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));

            IOrganizationServiceFactory servicefactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = servicefactory.CreateOrganizationService(context.UserId);

            //Build the following SQL query using QueryExpression:
            /* 
            select ID, Lead_Source, Start_Date, End_Date, Special_Processing_Rules, Lead_Routing_Type
            from
                Lead_Distribution a
                    inner join
                Lead_Routing_Type b
                    on a.Lead_Routing_Type = b.ID
                    inner join 
                Opportunity c
                    on b.ID = c.Routing_Type
            where 
            (
	            c.Lead_Source = "opportunity lead source" and c.Lead_Routing_Type = "opportunity of lead routing type"
            )
            */

            QueryExpression query = new QueryExpression()
            {
                Distinct = false,
                EntityName = pearl_leaddistribution.EntityLogicalName,
                ColumnSet = new ColumnSet("ID", "Lead_Source", "Start_Date", "End_Date", "Special_Processing_Rules", "Lead_Routing_Type"),
                LinkEntities = 
        {
            new LinkEntity 
            {
                JoinOperator = JoinOperator.Inner,
                LinkFromAttributeName = "Lead_Routing_Type",
                LinkFromEntityName = pearl_leaddistribution.EntityLogicalName,
                LinkToAttributeName = "ID",
                LinkToEntityName = pearl_leadroutingtype.EntityLogicalName
            },
            new LinkEntity
            {
                JoinOperator = JoinOperator.Inner,
                LinkFromAttributeName = "Routing_Type",
                LinkFromEntityName = pearl_leadroutingtype.EntityLogicalName,
                LinkToAttributeName = "ID",
                LinkToEntityName = Opportunity.EntityLogicalName,
            }
        },
                Criteria =
                {
                    Filters = 
            {
                new FilterExpression
                {
                     FilterOperator = LogicalOperator.And,
                    Conditions = 
                    {
                        new ConditionExpression("c.pearl_leadsource", ConditionOperator.Equal, sLeadSource),
                        new ConditionExpression("c.pearl_leadroutingtype", ConditionOperator.Equal, sLeadRoutingType)
                    },
                }
            }
                }
            };

            entityCollection = service.RetrieveMultiple(query).Entities;

            return entityCollection;
        }
    }
}
