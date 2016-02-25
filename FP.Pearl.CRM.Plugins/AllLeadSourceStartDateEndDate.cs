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
    public class AllLeadSourceStartDateEndDate
    {
        EntityCollection results = new EntityCollection();
        EntityCollection accResults = new EntityCollection();
        DataCollection<Entity> entityCollection = null;

        public DataCollection<Entity> Execute(IServiceProvider serviceProvider, string sLeadSource)
        {
            //Extract the tracing service for use in debugging sandboxed plug-ins.
            ITracingService tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));

            // Obtain the execution context from the service provider.
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));

            IOrganizationServiceFactory servicefactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = servicefactory.CreateOrganizationService(context.UserId);

            //Build the following SQL query using QueryExpression:
            /* 
            select ID, Name, Description
            from
                Lead_Source a
                    inner join
                Opportunity b
                    on a.ID = b.Lead_Source
            where 
            (
	            a.Lead_Source = "incoming string"
            )
            */

            QueryExpression query = new QueryExpression()
            {
                Distinct = false,
                EntityName = pearl_leadsource.EntityLogicalName,
                ColumnSet = new ColumnSet("ID", "Name", "Description"),
                LinkEntities = 
        {
            new LinkEntity 
            {
                JoinOperator = JoinOperator.Inner,
                LinkFromAttributeName = "ID",
                LinkFromEntityName = pearl_leadsource.EntityLogicalName,
                LinkToAttributeName = "Lead_Source",
                LinkToEntityName = Opportunity.EntityLogicalName
            }
        },
                Criteria =
                {
                    Filters = 
            {
                new FilterExpression
                {
                    Conditions = 
                    {
                        new ConditionExpression("Lead_Source", ConditionOperator.Equal, sLeadSource)
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
