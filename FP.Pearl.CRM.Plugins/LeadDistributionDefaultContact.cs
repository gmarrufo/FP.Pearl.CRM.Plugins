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
    public class LeadDistributionDefaultContact
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
            select ID, Default Distribution, Name, Default Contact
            from
                Lead Routing Type a
                    inner join
                Opportunity b
                    on a.id = b.routing type
            where 
            (
	            b.Lead Source = "incoming string"
            )
            */

            QueryExpression query = new QueryExpression()
            {
                Distinct = false,
                EntityName = pearl_leadroutingtype.EntityLogicalName,
                ColumnSet = new ColumnSet("ID", "Default Distribution", "Name", "Default Contact"),
                LinkEntities = 
        {
            new LinkEntity 
            {
                JoinOperator = JoinOperator.Inner,
                LinkFromAttributeName = "ID",
                LinkFromEntityName = pearl_leadroutingtype.EntityLogicalName,
                LinkToAttributeName = "Routing Type",
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
