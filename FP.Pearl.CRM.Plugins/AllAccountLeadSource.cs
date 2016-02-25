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
    public class AllAccountLeadSource
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
            select ID, Contacts, Name, Counties, Last Opportunity Assignment, Distribution Rules
            from
                Account a
                    inner join
                Lead_Routing_Type b
                    on a.Name = b.Name
                    inner join 
                Opportunity c
                    on b.ID = c.Routing_Type
            where 
            (
	            c.Lead_Source = "incoming string"
            )
            */

            QueryExpression query = new QueryExpression()
            {
                Distinct = false,
                EntityName = Account.EntityLogicalName,
                ColumnSet = new ColumnSet("ID", "Contacts", "Name", "Counties", "Last Opportunity Assignment", "Distribution Rules"),
                LinkEntities = 
        {
            new LinkEntity
            {
                JoinOperator = JoinOperator.Inner,
                LinkFromAttributeName = "pearl_DefaultAccount",
                LinkFromEntityName = Account.EntityLogicalName,
                LinkToAttributeName = "pearl_DefaultContact",
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
